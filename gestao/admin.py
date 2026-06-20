from django.contrib import admin
from .models import (
    Doador,
    CategoriaProduto,
    Produto,
    Campanha,
    Doacao,
    ItemDoacao,
    MovimentacaoEstoque,
    Beneficiario,
    Coleta,
    EntregaBeneficiario,
)


admin.site.register(Doador)
admin.site.register(CategoriaProduto)
admin.site.register(Produto)
admin.site.register(Campanha)
admin.site.register(Doacao)
admin.site.register(ItemDoacao)
admin.site.register(MovimentacaoEstoque)
admin.site.register(Beneficiario)
admin.site.register(Coleta)
admin.site.register(EntregaBeneficiario)