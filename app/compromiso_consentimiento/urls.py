from django.urls import path
from .views import index_compromiso_consentimiento

from .views import index_compromiso_consentimiento, get_redes_compromiso_consentimiento, RptconsentimientoCompromiso
from .views import get_microredes_compromiso_consentimiento, p_microredes_compromiso_consentimiento, RptPnPoblacionMicroRed
from .views import get_establecimientos_compromiso_consentimiento, p_microredes_establec_compromiso_consentimiento, p_establecimientos_compromiso_consentimiento, RptPnPoblacionEstablec
from .views import get_establecimientos_compromiso_consentimiento_h, p_microredes_establec_compromiso_consentimiento_h, p_establecimientos_compromiso_consentimiento_h, p_distritos_compromiso_consentimiento_h

from .views import get_provincias_compromiso_consentimiento,get_distritos_compromiso_consentimiento,p_distrito_compromiso_consentimiento

urlpatterns = [
    
    path('compromiso_consentimiento/', index_compromiso_consentimiento, name='index_compromiso_consentimiento'),

    ### BARRA HORIZONTAL
    
    path('get_establecimientos_compromiso_consentimiento_h/<int:establecimiento_id>/', get_establecimientos_compromiso_consentimiento_h, name='get_establecimientos_compromiso_consentimiento_h'),
    path('p_microredes_establec_compromiso_consentimiento_h/', p_microredes_establec_compromiso_consentimiento_h, name='p_microredes_establec_compromiso_consentimiento_h'),
    path('p_establecimiento_compromiso_consentimiento_h/', p_establecimientos_compromiso_consentimiento_h, name='p_establecimientos_compromiso_consentimiento_h'),
    path('p_distritos_compromiso_consentimiento_h/', p_distritos_compromiso_consentimiento_h, name='p_distritos_compromiso_consentimiento_h'),    
    
    ### SEGUIMIENTO NOMINAL
    ##-- AMBITO SALUD
    # redes
    path('get_redes_compromiso_consentimiento/<int:redes_id>/', get_redes_compromiso_consentimiento, name='get_redes_compromiso_consentimiento'),
    #-- redes excel
    path('rpt_compromiso_consentimiento_red_excel/', RptconsentimientoCompromiso.as_view(), name = 'rpt_compromiso_consentimiento_red_xls'),
    
    # microredes
    path('get_microredes_compromiso_consentimiento/<int:microredes_id>/', get_microredes_compromiso_consentimiento, name='get_microredes_compromiso_consentimiento'),
    path('p_microredes_compromiso_consentimiento/', p_microredes_compromiso_consentimiento, name='p_microredes_compromiso_consentimiento'),
    #-- microredes excel
    path('rpt_compromiso_consentimiento_microred_excel/', RptconsentimientoCompromiso.as_view(), name = 'rpt_compromiso_consentimiento_microred_xls'),
    
    # establecimientos
    path('get_establecimientos_compromiso_consentimiento/<int:establecimiento_id>/', get_establecimientos_compromiso_consentimiento, name='get_establecimientos_compromiso_consentimiento'),
    path('p_microredes_establec_compromiso_consentimiento/', p_microredes_establec_compromiso_consentimiento, name='p_microredes_establec_compromiso_consentimiento'),
    path('p_establecimiento_compromiso_consentimiento/', p_establecimientos_compromiso_consentimiento, name='p_establecimientos_compromiso_consentimiento'),       
    #-- estableccimiento excel
    path('rpt_compromiso_consentimiento_establec_excel/', RptconsentimientoCompromiso.as_view(), name = 'rpt_compromiso_consentimiento_establecimiento_xls'),
    
    ##-- AMBITO MUNICIPIO
    # provincia
    path('get_compromiso_consentimiento_provincia/<int:provincia_id>/', get_provincias_compromiso_consentimiento, name='get_provincias_compromiso_consentimiento'),
    #-- provincia excel
    path('rpt_compromiso_consentimiento_provincia_excel/', RptconsentimientoCompromiso.as_view(), name = 'rpt_compromiso_consentimiento_provincia_xls'),
    
    # distrito
    path('get_distrito_compromiso_consentimiento/<int:distrito_id>/', get_distritos_compromiso_consentimiento, name='get_distrito_compromiso_consentimiento'),
    path('p_distrito_compromiso_consentimiento/', p_distrito_compromiso_consentimiento, name='p_distrito_compromiso_consentimiento'),
    #-- distrito excel
    path('rpt_compromiso_consentimiento_distrito_excel/', RptconsentimientoCompromiso.as_view(), name = 'rpt_compromiso_consentimiento_distrito_xls'),
    
    
]