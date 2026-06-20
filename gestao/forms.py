from django import forms
from .models import Doador, Produto, Campanha, Doacao, Beneficiario, Coleta, MovimentacaoEstoque, EntregaBeneficiario, CategoriaProduto


class BaseForm(forms.ModelForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        for field in self.fields.values():
            field.widget.attrs.update({
                "class": "form-control"
            })


class DoadorForm(BaseForm):
    class Meta:
        model = Doador
        fields = "__all__"


class ProdutoForm(BaseForm):
    class Meta:
        model = Produto
        fields = "__all__"

class ProdutoForm(BaseForm):
    nova_categoria = forms.CharField(
        label="Categoria",
        max_length=100,
        required=True,
        widget=forms.TextInput(attrs={
            "class": "form-control",
            "placeholder": "Ex: Alimentos, Agasalhos, Higiene..."
        })
    )

    class Meta:
        model = Produto
        fields = ["nome", "nova_categoria", "unidade", "estoque_minimo", "ativo"]

    def save(self, commit=True):
        produto = super().save(commit=False)

        nome_categoria = self.cleaned_data["nova_categoria"].strip()

        categoria, criada = CategoriaProduto.objects.get_or_create(
            nome__iexact=nome_categoria,
            defaults={"nome": nome_categoria}
        )

        produto.categoria = categoria

        if commit:
            produto.save()

        return produto


class CampanhaForm(BaseForm):
    class Meta:
        model = Campanha
        fields = "__all__"
        widgets = {
            "data_inicio": forms.DateInput(attrs={"type": "date"}),
            "data_fim": forms.DateInput(attrs={"type": "date"}),
        }


class DoacaoForm(BaseForm):
    class Meta:
        model = Doacao
        fields = ["doador", "campanha", "data_recebimento", "observacao"]
        widgets = {
            "data_recebimento": forms.DateInput(attrs={"type": "date"}),
            "observacao": forms.Textarea(attrs={"rows": 3}),
        }


class BeneficiarioForm(BaseForm):
    class Meta:
        model = Beneficiario
        fields = "__all__"


class ColetaForm(BaseForm):
    class Meta:
        model = Coleta
        fields = "__all__"

class MovimentacaoEstoqueForm(BaseForm):
    class Meta:
        model = MovimentacaoEstoque
        fields = ["produto", "tipo", "quantidade", "observacao"]


class EntregaBeneficiarioForm(BaseForm):
    class Meta:
        model = EntregaBeneficiario
        fields = ["beneficiario", "produto", "quantidade", "data_entrega", "observacao"]
        widgets = {
            "data_entrega": forms.DateInput(attrs={"type": "date"}),
            "observacao": forms.Textarea(attrs={"rows": 3}),
        }


class BeneficiarioForm(BaseForm):
    class Meta:
        model = Beneficiario
        fields = "__all__"


class ColetaForm(BaseForm):
    class Meta:
        model = Coleta
        fields = "__all__"
        widgets = {
            "data_agendada": forms.DateTimeInput(attrs={"type": "datetime-local"}),
            "endereco_coleta": forms.Textarea(attrs={"rows": 3}),
            "observacao": forms.Textarea(attrs={"rows": 3}),
        }