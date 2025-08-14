from django.urls import path
from .views import index_compromiso_nino


from .views import get_provincias_compromiso, RptSituacionProvincia 
from .views import get_distritos_compromiso, p_distrito_compromiso, RptSituacionDistrito


urlpatterns = [
    
    path('compromiso_nino/', index_compromiso_nino, name='index_compromiso_nino'),

    # provincia
    path('get_compromiso_provincia/<int:provincia_id>/', get_provincias_compromiso, name='compromiso_get_provincias'),
    #-- provincia excel
    path('rpt_compromiso_provincia_excel/', RptSituacionProvincia.as_view(), name = 'rpt_compromiso_provincia_xls'),
    
    # distrito
    path('get_compromiso_distrito/<int:distrito_id>/', get_distritos_compromiso, name='compromiso_get_distrito'),
    path('p_distrito_compromiso/', p_distrito_compromiso, name='p_distrito_compromiso'),
    #-- distrito excel
    path('rpt_compromiso_distrito_excel/', RptSituacionDistrito.as_view(), name = 'rpt_compromiso_distrito_xls'),
    
]