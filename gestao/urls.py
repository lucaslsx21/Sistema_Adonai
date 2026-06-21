from django.urls import path
from . import views


urlpatterns = [
    path("", views.dashboard, name="dashboard"),

    # Doadores
    path("doadores/", views.doadores, name="doadores"),
    path("doadores/novo/", views.criar_doador, name="criar_doador"),
    path("doadores/editar/<int:pk>/", views.editar_doador, name="editar_doador"),
    path("doadores/excluir/<int:pk>/", views.excluir_doador, name="excluir_doador"),

    # Produtos
    path("produtos/", views.produtos, name="produtos"),
    path("produtos/novo/", views.criar_produto, name="criar_produto"),
    path("produtos/editar/<int:pk>/", views.editar_produto, name="editar_produto"),
    path("produtos/excluir/<int:pk>/", views.excluir_produto, name="excluir_produto"),

    # Campanhas
    path("campanhas/", views.campanhas, name="campanhas"),
    path("campanhas/novo/", views.criar_campanha, name="criar_campanha"),
    path("campanhas/editar/<int:pk>/", views.editar_campanha, name="editar_campanha"),
    path("campanhas/excluir/<int:pk>/", views.excluir_campanha, name="excluir_campanha"),

    # Doações
    path("doacoes/", views.doacoes, name="doacoes"),
    path("doacoes/novo/", views.criar_doacao, name="criar_doacao"),
    path("doacoes/editar/<int:pk>/", views.editar_doacao, name="editar_doacao"),
    path("doacoes/excluir/<int:pk>/", views.excluir_doacao, name="excluir_doacao"),

    # Estoque
    path("estoque/", views.estoque, name="estoque"),
    path("estoque/novo/", views.criar_movimentacao, name="criar_movimentacao"),
    path("estoque/excluir/<int:pk>/", views.excluir_movimentacao, name="excluir_movimentacao"),

    # Beneficiários
    path("beneficiarios/", views.beneficiarios, name="beneficiarios"),
    path("beneficiarios/novo/", views.criar_beneficiario, name="criar_beneficiario"),
    path("beneficiarios/editar/<int:pk>/", views.editar_beneficiario, name="editar_beneficiario"),
    path("beneficiarios/excluir/<int:pk>/", views.excluir_beneficiario, name="excluir_beneficiario"),

    # Coletas
    path("coletas/", views.coletas, name="coletas"),
    path("coletas/novo/", views.criar_coleta, name="criar_coleta"),
    path("coletas/editar/<int:pk>/", views.editar_coleta, name="editar_coleta"),
    path("coletas/excluir/<int:pk>/", views.excluir_coleta, name="excluir_coleta"),

    # Entregas
    path("entregas/", views.entregas, name="entregas"),
    path("entregas/novo/", views.criar_entrega, name="criar_entrega"),
    path("entregas/excluir/<int:pk>/", views.excluir_entrega, name="excluir_entrega"),

    # Relatórios
    path("relatorios/", views.relatorios, name="relatorios"),
    path("relatorio/<str:tipo>/excel/", views.exportar_excel, name="exportar_excel"),
    path("relatorio/<str:tipo>/word/", views.exportar_word, name="exportar_word"),
    path("relatorio/<str:tipo>/pdf/", views.exportar_pdf, name="exportar_pdf"),

    # Documentos / Importação
    path("documentos/", views.documentos, name="documentos"),
    path("importar-documentos/", views.importar_documentos, name="importar_documentos"),
]