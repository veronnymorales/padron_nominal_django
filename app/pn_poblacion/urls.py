from django.urls import path
from .views import index_pn_poblacion

from .views import index_pn_poblacion, get_redes_pn_poblacion, RptPnPoblacionRed
from .views import get_microredes_pn_poblacion, p_microredes_pn_poblacion, RptPnPoblacionMicroRed
from .views import get_establecimientos_pn_poblacion, p_microredes_establec_pn_poblacion, p_establecimientos_pn_poblacion, RptPnPoblacionEstablec
from .views import get_establecimientos_pn_poblacion_h, p_microredes_establec_pn_poblacion_h, p_establecimientos_pn_poblacion_h, p_distritos_pn_poblacion_h

from .views import get_provincias_pn_poblacion,get_distritos_pn_poblacion,p_distrito_pn_poblacion

urlpatterns = [
    
    path('pn_poblacion/', index_pn_poblacion, name='index_pn_poblacion'),

    ### BARRA HORIZONTAL
    
    path('get_establecimientos_pn_poblacion_h/<int:establecimiento_id>/', get_establecimientos_pn_poblacion_h, name='get_establecimientos_pn_poblacion_h'),
    path('p_microredes_establec_pn_poblacion_h/', p_microredes_establec_pn_poblacion_h, name='p_microredes_establec_pn_poblacion_h'),
    path('p_establecimiento_pn_poblacion_h/', p_establecimientos_pn_poblacion_h, name='p_establecimientos_pn_poblacion_h'),
    path('p_distritos_pn_poblacion_h/', p_distritos_pn_poblacion_h, name='p_distritos_pn_poblacion_h'),    
    
    ### SEGUIMIENTO NOMINAL
    ##-- AMBITO SALUD
    # redes
    path('get_redes_pn_poblacion/<int:redes_id>/', get_redes_pn_poblacion, name='get_redes_pn_poblacion'),
    #-- redes excel
    path('rpt_pn_poblacion_red_excel/', RptPnPoblacionRed.as_view(), name = 'rpt_pn_poblacion_red_xls'),
    
    # microredes
    path('get_microredes_pn_poblacion/<int:microredes_id>/', get_microredes_pn_poblacion, name='get_microredes_pn_poblacion'),
    path('p_microredes_pn_poblacion/', p_microredes_pn_poblacion, name='p_microredes_pn_poblacion'),
    #-- microredes excel
    path('rpt_pn_poblacion_microred_excel/', RptPnPoblacionMicroRed.as_view(), name = 'rpt_pn_poblacion_microred_xls'),
    
    # establecimientos
    path('get_establecimientos_pn_poblacion/<int:establecimiento_id>/', get_establecimientos_pn_poblacion, name='get_establecimientos_pn_poblacion'),
    path('p_microredes_establec_pn_poblacion/', p_microredes_establec_pn_poblacion, name='p_microredes_establec_pn_poblacion'),
    path('p_establecimiento_pn_poblacion/', p_establecimientos_pn_poblacion, name='p_establecimientos_pn_poblacion'),       
    #-- estableccimiento excel
    path('rpt_pn_poblacion_establec_excel/', RptPnPoblacionEstablec.as_view(), name = 'rpt_pn_poblacion_establecimiento_xls'),
    
    ##-- AMBITO MUNICIPIO
    # provincia
    path('get_pn_poblacion_provincia/<int:provincia_id>/', get_provincias_pn_poblacion, name='get_provincias_pn_poblacion'),
    #-- provincia excel
    #path('rpt_pn_poblacion_provincia_excel/', RptPnPoblacionProvincia.as_view(), name = 'rpt_pn_poblacion_provincia_xls'),
    
    # distrito
    path('get_distrito_pn_poblacion/<int:distrito_id>/', get_distritos_pn_poblacion, name='get_distrito_pn_poblacion'),
    path('p_distrito_pn_poblacion/', p_distrito_pn_poblacion, name='p_distrito_pn_poblacion'),
    #-- distrito excel
    #path('rpt_compromiso_distrito_excel/', RptSituacionDistrito.as_view(), name = 'rpt_compromiso_distrito_xls'),
    
    
]