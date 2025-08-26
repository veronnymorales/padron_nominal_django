from django.urls import path
from .views import index_paquete_compromiso

from .views import index_paquete_compromiso, get_redes_paquete_compromiso, RptPnPoblacionRed
from .views import get_microredes_paquete_compromiso, p_microredes_paquete_compromiso, RptPnPoblacionMicroRed
from .views import get_establecimientos_paquete_compromiso, p_microredes_establec_paquete_compromiso, p_establecimientos_paquete_compromiso, RptPnPoblacionEstablec
from .views import get_establecimientos_paquete_compromiso_h, p_microredes_establec_paquete_compromiso_h, p_establecimientos_paquete_compromiso_h, p_distritos_paquete_compromiso_h

from .views import get_provincias_paquete_compromiso,get_distritos_paquete_compromiso,p_distrito_paquete_compromiso

urlpatterns = [
    
    path('paquete_compromiso/', index_paquete_compromiso, name='index_paquete_compromiso'),

    ### BARRA HORIZONTAL
    
    path('get_establecimientos_paquete_compromiso_h/<int:establecimiento_id>/', get_establecimientos_paquete_compromiso_h, name='get_establecimientos_paquete_compromiso_h'),
    path('p_microredes_establec_paquete_compromiso_h/', p_microredes_establec_paquete_compromiso_h, name='p_microredes_establec_paquete_compromiso_h'),
    path('p_establecimiento_paquete_compromiso_h/', p_establecimientos_paquete_compromiso_h, name='p_establecimientos_paquete_compromiso_h'),
    path('p_distritos_paquete_compromiso_h/', p_distritos_paquete_compromiso_h, name='p_distritos_paquete_compromiso_h'),    
    
    ### SEGUIMIENTO NOMINAL
    ##-- AMBITO SALUD
    # redes
    path('get_redes_paquete_compromiso/<int:redes_id>/', get_redes_paquete_compromiso, name='get_redes_paquete_compromiso'),
    #-- redes excel
    path('rpt_paquete_compromiso_red_excel/', RptPnPoblacionRed.as_view(), name = 'rpt_paquete_compromiso_red_xls'),
    
    # microredes
    path('get_microredes_paquete_compromiso/<int:microredes_id>/', get_microredes_paquete_compromiso, name='get_microredes_paquete_compromiso'),
    path('p_microredes_paquete_compromiso/', p_microredes_paquete_compromiso, name='p_microredes_paquete_compromiso'),
    #-- microredes excel
    path('rpt_paquete_compromiso_microred_excel/', RptPnPoblacionMicroRed.as_view(), name = 'rpt_paquete_compromiso_microred_xls'),
    
    # establecimientos
    path('get_establecimientos_paquete_compromiso/<int:establecimiento_id>/', get_establecimientos_paquete_compromiso, name='get_establecimientos_paquete_compromiso'),
    path('p_microredes_establec_paquete_compromiso/', p_microredes_establec_paquete_compromiso, name='p_microredes_establec_paquete_compromiso'),
    path('p_establecimiento_paquete_compromiso/', p_establecimientos_paquete_compromiso, name='p_establecimientos_paquete_compromiso'),       
    #-- estableccimiento excel
    path('rpt_paquete_compromiso_establec_excel/', RptPnPoblacionEstablec.as_view(), name = 'rpt_paquete_compromiso_establecimiento_xls'),
    
    ##-- AMBITO MUNICIPIO
    # provincia
    path('get_paquete_compromiso_provincia/<int:provincia_id>/', get_provincias_paquete_compromiso, name='get_provincias_paquete_compromiso'),
    #-- provincia excel
    #path('rpt_paquete_compromiso_provincia_excel/', RptPnPoblacionProvincia.as_view(), name = 'rpt_paquete_compromiso_provincia_xls'),
    
    # distrito
    path('get_distrito_paquete_compromiso/<int:distrito_id>/', get_distritos_paquete_compromiso, name='get_distrito_paquete_compromiso'),
    path('p_distrito_paquete_compromiso/', p_distrito_paquete_compromiso, name='p_distrito_paquete_compromiso'),
    #-- distrito excel
    #path('rpt_compromiso_distrito_excel/', RptSituacionDistrito.as_view(), name = 'rpt_compromiso_distrito_xls'),
    
    
]