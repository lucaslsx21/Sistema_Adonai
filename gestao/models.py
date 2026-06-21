from django.db import models
from django.contrib.auth.models import User


class Doador(models.Model):
    TIPO_CHOICES = [
        ("PF", "Pessoa Física"),
        ("PJ", "Pessoa Jurídica"),
    ]

    nome = models.CharField(max_length=200)
    tipo = models.CharField(max_length=2, choices=TIPO_CHOICES)
    telefone = models.CharField(max_length=20)
    email = models.EmailField(blank=True, null=True)
    endereco = models.CharField(max_length=255, blank=True)
    cidade = models.CharField(max_length=100, blank=True)
    ativo = models.BooleanField(default=True)
    criado_em = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.nome


class CategoriaProduto(models.Model):
    nome = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.nome


class Produto(models.Model):
    UNIDADES = [
        ("UN", "Unidade"),
        ("KG", "Quilo"),
        ("L", "Litro"),
        ("CX", "Caixa"),
        ("PCT", "Pacote"),
    ]

    nome = models.CharField(max_length=200)
    categoria = models.ForeignKey(CategoriaProduto, on_delete=models.PROTECT)
    unidade = models.CharField(max_length=10, choices=UNIDADES)
    estoque_minimo = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    ativo = models.BooleanField(default=True)

    def estoque_atual(self):
        entradas = self.movimentacoes.filter(tipo="ENTRADA").aggregate(
            total=models.Sum("quantidade")
        )["total"] or 0

        saidas = self.movimentacoes.filter(tipo="SAIDA").aggregate(
            total=models.Sum("quantidade")
        )["total"] or 0

        return entradas - saidas

    def tem_estoque_suficiente(self, quantidade):
        return self.estoque_atual() >= quantidade

    def __str__(self):
        return self.nome


class Campanha(models.Model):
    nome = models.CharField(max_length=200)
    descricao = models.TextField(blank=True)
    data_inicio = models.DateField()
    data_fim = models.DateField()
    ativa = models.BooleanField(default=True)

    def __str__(self):
        return self.nome


class Doacao(models.Model):
    doador = models.ForeignKey(Doador, on_delete=models.PROTECT)
    campanha = models.ForeignKey(
        Campanha,
        on_delete=models.SET_NULL,
        null=True,
        blank=True
    )
    data_recebimento = models.DateField()
    observacao = models.TextField(blank=True)
    criado_por = models.ForeignKey(User, on_delete=models.PROTECT)
    criado_em = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Doação de {self.doador.nome}"


class ItemDoacao(models.Model):
    doacao = models.ForeignKey(
        Doacao,
        related_name="itens",
        on_delete=models.CASCADE
    )
    produto = models.ForeignKey(Produto, on_delete=models.PROTECT)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2)
    validade = models.DateField(null=True, blank=True)

    def __str__(self):
        return f"{self.produto.nome} - {self.quantidade}"


class MovimentacaoEstoque(models.Model):
    TIPOS = [
        ("ENTRADA", "Entrada"),
        ("SAIDA", "Saída"),
    ]

    produto = models.ForeignKey(
        Produto,
        related_name="movimentacoes",
        on_delete=models.PROTECT
    )
    tipo = models.CharField(max_length=10, choices=TIPOS)
    quantidade = models.DecimalField(max_digits=10, decimal_places=2)
    data = models.DateTimeField(auto_now_add=True)
    observacao = models.TextField(blank=True)

    def __str__(self):
        return f"{self.tipo} - {self.produto.nome}"


class Beneficiario(models.Model):
    nome = models.CharField(max_length=200)
    cpf = models.CharField(max_length=14, blank=True)
    telefone = models.CharField(max_length=20, blank=True)
    endereco = models.TextField(blank=True)
    quantidade_pessoas = models.IntegerField(default=1)
    observacoes = models.TextField(blank=True)
    ativo = models.BooleanField(default=True)

    def __str__(self):
        return self.nome


class Coleta(models.Model):
    STATUS = [
        ("AGENDADA", "Agendada"),
        ("EM_ROTA", "Em rota"),
        ("CONCLUIDA", "Concluída"),
        ("CANCELADA", "Cancelada"),
    ]

    doador = models.ForeignKey(Doador, on_delete=models.PROTECT)
    data_agendada = models.DateTimeField()
    endereco_coleta = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS, default="AGENDADA")
    observacao = models.TextField(blank=True)

    def __str__(self):
        return f"Coleta - {self.doador.nome}"


class EntregaBeneficiario(models.Model):
    beneficiario = models.ForeignKey(
        Beneficiario,
        on_delete=models.PROTECT
    )
    produto = models.ForeignKey(
        Produto,
        on_delete=models.PROTECT
    )
    quantidade = models.DecimalField(
        max_digits=10,
        decimal_places=2
    )
    data_entrega = models.DateField()
    observacao = models.TextField(blank=True)
    criado_em = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.beneficiario.nome} - {self.produto.nome}"


class DocumentoImportado(models.Model):
    TIPO_ARQUIVO_CHOICES = [
        ("EXCEL", "Excel"),
        ("WORD", "Word"),
        ("PDF", "PDF"),
        ("OUTRO", "Outro"),
    ]

    TIPO_IMPORTACAO_CHOICES = [
        ("PRODUTOS", "Produtos"),
        ("DOADORES", "Doadores"),
        ("BENEFICIARIOS", "Beneficiários"),
        ("DOACOES", "Doações"),
        ("ESTOQUE", "Estoque"),
        ("CAMPANHAS", "Campanhas"),
    ]

    titulo = models.CharField(max_length=150)
    arquivo = models.FileField(upload_to="documentos_importados/")
    tipo = models.CharField(
        max_length=20,
        choices=TIPO_ARQUIVO_CHOICES,
        default="OUTRO"
    )
    tipo_importacao = models.CharField(
        max_length=30,
        choices=TIPO_IMPORTACAO_CHOICES,
        default="PRODUTOS"
    )
    texto_extraido = models.TextField(blank=True, null=True)
    importado_em = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.titulo


class RegistroImportado(models.Model):
    documento = models.ForeignKey(
        DocumentoImportado,
        on_delete=models.CASCADE,
        related_name="registros"
    )
    conteudo = models.TextField()
    criado_em = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.conteudo[:80]