from django.urls import path
from .views import index_situacion_padron
from .views import get_provincias_situacion, RptSituacionProvincia 
from .views import get_distritos_situacion, p_distritos_situacion, RptSituacionDistrito



urlpatterns = [
    
    
    path('situacion_padron/', index_situacion_padron, name='index_situacion_padron'),
    
    ### SEGUIMIENTO
    # provincia
    path('get_provincia_situacion/<int:provincia_id>/', get_provincias_situacion, name='pn_situacion_actual_get_provincias'),
    #-- provincia excel
    path('rpt_situacion_provincia_excel/', RptSituacionProvincia.as_view(), name = 'rpt_situacion_provincia_xls'),
    
    # distrito
    path('get_distrito_situacion/<int:distrito_id>/', get_distritos_situacion, name='pn_situacion_actual_get_distritos'),
    path('p_distrito_situacion/', p_distritos_situacion, name='p_distrito_situacion'),
    #-- distrito excel
    path('rpt_situacion_distrito_excel/', RptSituacionDistrito.as_view(), name = 'rpt_situacion_distrito_xls'),
    
    
    
]