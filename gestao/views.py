from django.db.models import Count, Sum
from django.db.models.functions import TruncMonth
import json
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.contrib import messages

from openpyxl import Workbook
from docx import Document
from reportlab.pdfgen import canvas

from .models import Doador, Produto, Campanha, Doacao, Beneficiario, Coleta, MovimentacaoEstoque, EntregaBeneficiario
from .forms import DoadorForm, ProdutoForm, CampanhaForm, DoacaoForm, BeneficiarioForm, ColetaForm, MovimentacaoEstoqueForm, EntregaBeneficiarioForm


@login_required
def dashboard(request):
    doacoes_por_mes = (
        Doacao.objects
        .annotate(mes=TruncMonth("data_recebimento"))
        .values("mes")
        .annotate(total=Count("id"))
        .order_by("mes")
    )

    campanhas = (
        Campanha.objects
        .annotate(total_doacoes=Count("doacao"))
        .order_by("-total_doacoes")[:5]
    )

    produtos_estoque = Produto.objects.all()

    produtos_baixo_estoque = [
        produto for produto in produtos_estoque
        if produto.estoque_atual() <= produto.estoque_minimo
    ]

    context = {
        "total_doadores": Doador.objects.count(),
        "total_produtos": Produto.objects.count(),
        "total_doacoes": Doacao.objects.count(),
        "total_campanhas": Campanha.objects.filter(ativa=True).count(),
        "total_beneficiarios": Beneficiario.objects.count(),
        "produtos_baixo_estoque": produtos_baixo_estoque,
        "ultimas_doacoes": Doacao.objects.order_by("-criado_em")[:5],

        "grafico_meses": json.dumps([
            item["mes"].strftime("%m/%Y") for item in doacoes_por_mes
        ]),

        "grafico_doacoes": json.dumps([
            item["total"] for item in doacoes_por_mes
        ]),

        "grafico_campanhas_labels": json.dumps([
            item.nome for item in campanhas
        ]),

        "grafico_campanhas_valores": json.dumps([
            item.total_doacoes for item in campanhas
        ]),
    }

    return render(request, "gestao/dashboard.html", context)


# =========================
# CRUD GENÉRICO
# =========================

def crud_list(request, model, template, context_name):
    busca = request.GET.get("busca")

    if busca:
        dados = model.objects.filter(nome__icontains=busca)
    else:
        dados = model.objects.all()

    return render(request, template, {context_name: dados})


@login_required
def doadores(request):
    return crud_list(request, Doador, "gestao/doadores.html", "dados")


@login_required
def produtos(request):
    return crud_list(request, Produto, "gestao/produtos.html", "dados")


@login_required
def campanhas(request):
    return crud_list(request, Campanha, "gestao/campanhas.html", "dados")


@login_required
def doacoes(request):
    dados = Doacao.objects.all().order_by("-data_recebimento")
    return render(request, "gestao/doacoes.html", {"dados": dados})


# =========================
# FORMULÁRIOS
# =========================

@login_required
def criar_doador(request):
    form = DoadorForm(request.POST or None)

    if form.is_valid():
        form.save()
        messages.success(request, "Doador cadastrado com sucesso.")
        return redirect("doadores")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Novo Doador"
    })


@login_required
def editar_doador(request, pk):
    item = get_object_or_404(Doador, pk=pk)
    form = DoadorForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Doador atualizado com sucesso.")
        return redirect("doadores")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Doador"
    })


@login_required
def excluir_doador(request, pk):
    item = get_object_or_404(Doador, pk=pk)
    item.delete()
    messages.success(request, "Doador excluído com sucesso.")
    return redirect("doadores")


@login_required
def criar_produto(request):
    form = ProdutoForm(request.POST or None)

    if form.is_valid():
        form.save()
        messages.success(request, "Produto cadastrado com sucesso.")
        return redirect("produtos")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Novo Produto"
    })


@login_required
def editar_produto(request, pk):
    item = get_object_or_404(Produto, pk=pk)
    form = ProdutoForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Produto atualizado com sucesso.")
        return redirect("produtos")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Produto"
    })


@login_required
def excluir_produto(request, pk):
    item = get_object_or_404(Produto, pk=pk)
    item.delete()
    messages.success(request, "Produto excluído com sucesso.")
    return redirect("produtos")


@login_required
def criar_campanha(request):
    form = CampanhaForm(request.POST or None)

    if form.is_valid():
        form.save()
        messages.success(request, "Campanha cadastrada com sucesso.")
        return redirect("campanhas")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Nova Campanha"
    })


@login_required
def editar_campanha(request, pk):
    item = get_object_or_404(Campanha, pk=pk)
    form = CampanhaForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Campanha atualizada com sucesso.")
        return redirect("campanhas")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Campanha"
    })


@login_required
def excluir_campanha(request, pk):
    item = get_object_or_404(Campanha, pk=pk)
    item.delete()
    messages.success(request, "Campanha excluída com sucesso.")
    return redirect("campanhas")


@login_required
def criar_doacao(request):
    form = DoacaoForm(request.POST or None)

    if form.is_valid():
        doacao = form.save(commit=False)
        doacao.criado_por = request.user
        doacao.save()

        messages.success(request, "Doação cadastrada com sucesso.")
        return redirect("doacoes")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Nova Doação"
    })


@login_required
def editar_doacao(request, pk):
    item = get_object_or_404(Doacao, pk=pk)
    form = DoacaoForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Doação atualizada com sucesso.")
        return redirect("doacoes")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Doação"
    })


@login_required
def excluir_doacao(request, pk):
    item = get_object_or_404(Doacao, pk=pk)
    item.delete()
    messages.success(request, "Doação excluída com sucesso.")
    return redirect("doacoes")


# =========================
# RELATÓRIOS
# =========================

def get_model_data(tipo, request=None):
    data_inicio = request.GET.get("data_inicio") if request else None
    data_fim = request.GET.get("data_fim") if request else None
    campanha_id = request.GET.get("campanha") if request else None
    produto_id = request.GET.get("produto") if request else None
    doador_id = request.GET.get("doador") if request else None

    if tipo == "doadores":
        dados = Doador.objects.all()
        return dados, ["Nome", "Tipo", "Telefone", "Email", "Cidade"]

    if tipo == "produtos":
        dados = Produto.objects.all()
        return dados, ["Nome", "Categoria", "Unidade", "Estoque mínimo"]

    if tipo == "campanhas":
        dados = Campanha.objects.all()

        if data_inicio:
            dados = dados.filter(data_inicio__gte=data_inicio)

        if data_fim:
            dados = dados.filter(data_fim__lte=data_fim)

        return dados, ["Nome", "Início", "Fim", "Ativa"]

    if tipo == "doacoes":
        dados = Doacao.objects.all()

        if data_inicio:
            dados = dados.filter(data_recebimento__gte=data_inicio)

        if data_fim:
            dados = dados.filter(data_recebimento__lte=data_fim)

        if campanha_id:
            dados = dados.filter(campanha_id=campanha_id)

        if doador_id:
            dados = dados.filter(doador_id=doador_id)

        return dados, ["Doador", "Campanha", "Data", "Observação"]

    if tipo == "beneficiarios":
        dados = Beneficiario.objects.all()
        return dados, ["Nome", "Telefone", "Pessoas", "Ativo"]

    if tipo == "coletas":
        dados = Coleta.objects.all()

        if data_inicio:
            dados = dados.filter(data_agendada__date__gte=data_inicio)

        if data_fim:
            dados = dados.filter(data_agendada__date__lte=data_fim)

        if doador_id:
            dados = dados.filter(doador_id=doador_id)

        return dados, ["Doador", "Data", "Status", "Endereço"]

    if tipo == "estoque":
        dados = MovimentacaoEstoque.objects.all()

        if data_inicio:
            dados = dados.filter(data__date__gte=data_inicio)

        if data_fim:
            dados = dados.filter(data__date__lte=data_fim)

        if produto_id:
            dados = dados.filter(produto_id=produto_id)

        return dados, ["Produto", "Tipo", "Quantidade", "Data"]

    return [], []

def linha_objeto(tipo, obj):
    if tipo == "doadores":
        return [
            obj.nome,
            obj.get_tipo_display(),
            obj.telefone,
            obj.email,
            obj.cidade,
        ]

    if tipo == "produtos":
        return [
            obj.nome,
            obj.categoria.nome,
            obj.get_unidade_display(),
            obj.estoque_minimo,
        ]

    if tipo == "campanhas":
        return [
            obj.nome,
            obj.data_inicio,
            obj.data_fim,
            "Sim" if obj.ativa else "Não",
        ]

    if tipo == "doacoes":
        return [
            obj.doador.nome,
            obj.campanha.nome if obj.campanha else "Sem campanha",
            obj.data_recebimento,
            obj.observacao,
        ]
    if tipo == "beneficiarios":
        return [
            obj.nome,
            obj.telefone,
            obj.quantidade_pessoas,
            "Sim" if obj.ativo else "Não",
        ]

    if tipo == "coletas":
        return [
            obj.doador.nome,
            obj.data_agendada,
            obj.get_status_display(),
            obj.endereco_coleta,
        ]

    if tipo == "estoque":
        return [
            obj.produto.nome,
            obj.get_tipo_display(),
            obj.quantidade,
            obj.data,
        ]


@login_required
def exportar_excel(request, tipo):
    dados, colunas = get_model_data(tipo, request)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Relatório {tipo}"

    ws.append(colunas)

    for obj in dados:
        ws.append(linha_objeto(tipo, obj))

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    response["Content-Disposition"] = f'attachment; filename="relatorio_{tipo}.xlsx"'

    wb.save(response)
    return response


@login_required
def exportar_word(request, tipo):
    dados, colunas = get_model_data(tipo, request)

    doc = Document()
    doc.add_heading(f"Relatório de {tipo.title()}", 0)

    table = doc.add_table(rows=1, cols=len(colunas))
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells

    for i, coluna in enumerate(colunas):
        hdr_cells[i].text = coluna

    for obj in dados:
        row_cells = table.add_row().cells
        linha = linha_objeto(tipo, obj)

        for i, valor in enumerate(linha):
            row_cells[i].text = str(valor)

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    response["Content-Disposition"] = f'attachment; filename="relatorio_{tipo}.docx"'

    doc.save(response)
    return response


@login_required
def exportar_pdf(request, tipo):
    dados, colunas = get_model_data(tipo, request)

    response = HttpResponse(content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="relatorio_{tipo}.pdf"'

    p = canvas.Canvas(response)
    p.setFont("Helvetica-Bold", 16)
    p.drawString(50, 800, f"Relatório de {tipo.title()}")

    y = 760
    p.setFont("Helvetica-Bold", 9)

    x = 50
    for coluna in colunas:
        p.drawString(x, y, str(coluna))
        x += 120

    y -= 20
    p.setFont("Helvetica", 8)

    for obj in dados:
        x = 50
        linha = linha_objeto(tipo, obj)

        for valor in linha:
            p.drawString(x, y, str(valor)[:18])
            x += 120

        y -= 20

        if y < 50:
            p.showPage()
            y = 800

    p.save()
    return response

@login_required
def estoque(request):
    dados = MovimentacaoEstoque.objects.all().order_by("-data")
    produtos = Produto.objects.all().order_by("nome")

    return render(request, "gestao/estoque.html", {
        "dados": dados,
        "produtos": produtos,
    })


@login_required
def criar_movimentacao(request):
    form = MovimentacaoEstoqueForm(request.POST or None)

    if form.is_valid():
        movimentacao = form.save(commit=False)

        if movimentacao.tipo == "SAIDA":
            if not movimentacao.produto.tem_estoque_suficiente(movimentacao.quantidade):
                messages.error(
                    request,
                    f"Estoque insuficiente. Estoque atual: {movimentacao.produto.estoque_atual()}"
                )
                return render(request, "gestao/form.html", {
                    "form": form,
                    "titulo": "Nova Movimentação de Estoque"
                })

        movimentacao.save()
        messages.success(request, "Movimentação registrada com sucesso.")
        return redirect("estoque")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Nova Movimentação de Estoque"
    })

@login_required
def excluir_movimentacao(request, pk):
    item = get_object_or_404(MovimentacaoEstoque, pk=pk)
    item.delete()
    messages.success(request, "Movimentação excluída com sucesso.")
    return redirect("estoque")


@login_required
def beneficiarios(request):
    busca = request.GET.get("busca")

    if busca:
        dados = Beneficiario.objects.filter(nome__icontains=busca)
    else:
        dados = Beneficiario.objects.all().order_by("nome")

    return render(request, "gestao/beneficiarios.html", {"dados": dados})


@login_required
def criar_beneficiario(request):
    form = BeneficiarioForm(request.POST or None)

    if form.is_valid():
        form.save()
        messages.success(request, "Beneficiário cadastrado com sucesso.")
        return redirect("beneficiarios")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Novo Beneficiário"
    })


@login_required
def editar_beneficiario(request, pk):
    item = get_object_or_404(Beneficiario, pk=pk)
    form = BeneficiarioForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Beneficiário atualizado com sucesso.")
        return redirect("beneficiarios")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Beneficiário"
    })


@login_required
def excluir_beneficiario(request, pk):
    item = get_object_or_404(Beneficiario, pk=pk)
    item.delete()
    messages.success(request, "Beneficiário excluído com sucesso.")
    return redirect("beneficiarios")


@login_required
def coletas(request):
    dados = Coleta.objects.all().order_by("-data_agendada")
    return render(request, "gestao/coletas.html", {"dados": dados})


@login_required
def criar_coleta(request):
    form = ColetaForm(request.POST or None)

    if form.is_valid():
        form.save()
        messages.success(request, "Coleta cadastrada com sucesso.")
        return redirect("coletas")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Nova Coleta"
    })


@login_required
def editar_coleta(request, pk):
    item = get_object_or_404(Coleta, pk=pk)
    form = ColetaForm(request.POST or None, instance=item)

    if form.is_valid():
        form.save()
        messages.success(request, "Coleta atualizada com sucesso.")
        return redirect("coletas")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Editar Coleta"
    })


@login_required
def excluir_coleta(request, pk):
    item = get_object_or_404(Coleta, pk=pk)
    item.delete()
    messages.success(request, "Coleta excluída com sucesso.")
    return redirect("coletas")


@login_required
def entregas(request):
    dados = EntregaBeneficiario.objects.all().order_by("-data_entrega")
    return render(request, "gestao/entregas.html", {"dados": dados})


@login_required
def criar_entrega(request):
    form = EntregaBeneficiarioForm(request.POST or None)

    if form.is_valid():
        entrega = form.save(commit=False)

        if not entrega.produto.tem_estoque_suficiente(entrega.quantidade):
            messages.error(
                request,
                f"Estoque insuficiente para entregar {entrega.produto.nome}. Estoque atual: {entrega.produto.estoque_atual()}"
            )
            return render(request, "gestao/form.html", {
                "form": form,
                "titulo": "Nova Entrega para Beneficiário"
            })

        entrega.save()

        MovimentacaoEstoque.objects.create(
            produto=entrega.produto,
            tipo="SAIDA",
            quantidade=entrega.quantidade,
            observacao=f"Entrega para beneficiário: {entrega.beneficiario.nome}"
        )

        messages.success(request, "Entrega cadastrada e estoque atualizado.")
        return redirect("entregas")

    return render(request, "gestao/form.html", {
        "form": form,
        "titulo": "Nova Entrega para Beneficiário"
    })


@login_required
def excluir_entrega(request, pk):
    item = get_object_or_404(EntregaBeneficiario, pk=pk)
    item.delete()
    messages.success(request, "Entrega excluída com sucesso.")
    return redirect("entregas")


@login_required
def relatorios(request):
    context = {
        "doadores": Doador.objects.all().order_by("nome"),
        "campanhas": Campanha.objects.all().order_by("nome"),
        "produtos": Produto.objects.all().order_by("nome"),
    }

    return render(request, "gestao/relatorios.html", context)