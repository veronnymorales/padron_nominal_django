from django.urls import path
from .views import index_compromiso_indicador

from .views import index_compromiso_indicador, get_redes_compromiso_indicador, RptIndicadorCompromiso
from .views import get_microredes_compromiso_indicador, p_microredes_compromiso_indicador, RptPnPoblacionMicroRed
from .views import get_establecimientos_compromiso_indicador, p_microredes_establec_compromiso_indicador, p_establecimientos_compromiso_indicador, RptPnPoblacionEstablec
from .views import get_establecimientos_compromiso_indicador_h, p_microredes_establec_compromiso_indicador_h, p_establecimientos_compromiso_indicador_h, p_distritos_compromiso_indicador_h

from .views import get_provincias_compromiso_indicador,get_distritos_compromiso_indicador,p_distrito_compromiso_indicador

urlpatterns = [
    
    path('compromiso_indicador/', index_compromiso_indicador, name='index_compromiso_indicador'),

    ### BARRA HORIZONTAL
    
    path('get_establecimientos_compromiso_indicador_h/<int:establecimiento_id>/', get_establecimientos_compromiso_indicador_h, name='get_establecimientos_compromiso_indicador_h'),
    path('p_microredes_establec_compromiso_indicador_h/', p_microredes_establec_compromiso_indicador_h, name='p_microredes_establec_compromiso_indicador_h'),
    path('p_establecimiento_compromiso_indicador_h/', p_establecimientos_compromiso_indicador_h, name='p_establecimientos_compromiso_indicador_h'),
    path('p_distritos_compromiso_indicador_h/', p_distritos_compromiso_indicador_h, name='p_distritos_compromiso_indicador_h'),    
    
    ### SEGUIMIENTO NOMINAL
    ##-- AMBITO SALUD
    # redes
    path('get_redes_compromiso_indicador/<int:redes_id>/', get_redes_compromiso_indicador, name='get_redes_compromiso_indicador'),
    #-- redes excel
    path('rpt_compromiso_indicador_red_excel/', RptIndicadorCompromiso.as_view(), name = 'rpt_compromiso_indicador_red_xls'),
    
    # microredes
    path('get_microredes_compromiso_indicador/<int:microredes_id>/', get_microredes_compromiso_indicador, name='get_microredes_compromiso_indicador'),
    path('p_microredes_compromiso_indicador/', p_microredes_compromiso_indicador, name='p_microredes_compromiso_indicador'),
    #-- microredes excel
    path('rpt_compromiso_indicador_microred_excel/', RptIndicadorCompromiso.as_view(), name = 'rpt_compromiso_indicador_microred_xls'),
    
    # establecimientos
    path('get_establecimientos_compromiso_indicador/<int:establecimiento_id>/', get_establecimientos_compromiso_indicador, name='get_establecimientos_compromiso_indicador'),
    path('p_microredes_establec_compromiso_indicador/', p_microredes_establec_compromiso_indicador, name='p_microredes_establec_compromiso_indicador'),
    path('p_establecimiento_compromiso_indicador/', p_establecimientos_compromiso_indicador, name='p_establecimientos_compromiso_indicador'),       
    #-- estableccimiento excel
    path('rpt_compromiso_indicador_establec_excel/', RptIndicadorCompromiso.as_view(), name = 'rpt_compromiso_indicador_establecimiento_xls'),
    
    ##-- AMBITO MUNICIPIO
    # provincia
    path('get_compromiso_indicador_provincia/<int:provincia_id>/', get_provincias_compromiso_indicador, name='get_provincias_compromiso_indicador'),
    #-- provincia excel
    path('rpt_compromiso_indicador_provincia_excel/', RptIndicadorCompromiso.as_view(), name = 'rpt_compromiso_indicador_provincia_xls'),
    
    # distrito
    path('get_distrito_compromiso_indicador/<int:distrito_id>/', get_distritos_compromiso_indicador, name='get_distrito_compromiso_indicador'),
    path('p_distrito_compromiso_indicador/', p_distrito_compromiso_indicador, name='p_distrito_compromiso_indicador'),
    #-- distrito excel
    path('rpt_compromiso_indicador_distrito_excel/', RptIndicadorCompromiso.as_view(), name = 'rpt_compromiso_indicador_distrito_xls'),
    
    
]