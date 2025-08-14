from django.urls import path
from .views import index_pacto_nino


from .views import get_provincias_pacto, RptPactoProvincia 
from .views import get_distritos_pacto, p_distrito_pacto, RptPactoDistrito


urlpatterns = [
    
    path('pacto_nino/', index_pacto_nino, name='index_pacto_nino'),

    # provincia
    path('get_pacto_provincia/<int:provincia_id>/', get_provincias_pacto, name='pacto_get_provincias'),
    #-- provincia excel
    path('rpt_pacto_provincia_excel/', RptPactoProvincia.as_view(), name = 'rpt_pacto_provincia_xls'),
    
    # distrito
    path('get_pacto_distrito/<int:distrito_id>/', get_distritos_pacto, name='pacto_get_distrito'),
    path('p_distrito_pacto/', p_distrito_pacto, name='p_distrito_pacto'),
    #-- distrito excel
    path('rpt_pacto_distrito_excel/', RptPactoDistrito.as_view(), name = 'rpt_pacto_distrito_xls'),
    
]