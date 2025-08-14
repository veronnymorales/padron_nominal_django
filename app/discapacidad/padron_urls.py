from django.urls import path
from .padron_views import DirectorioMunicipioCreateView, DirectorioMunicipioListView, directorio_municipalidad_detail,DirectorioMunicipioListViewPublic
from .padron_views import DirectorioSaludCreateView, DirectorioSaludListView, directorio_salud_detail,DirectorioSaludListViewPublic
from .padron_views import index_sello, sello_get_provincias, sello_get_distritos, sello_p_distritos
from .padron_views import RptOperacinalDist


urlpatterns = [

    # -- DIRECTORIO MUNICIPIO
    path('municipio/create/', DirectorioMunicipioCreateView.as_view(), name='municipio-create'),
    
    path('municipio/list/', DirectorioMunicipioListView.as_view(), name='municipio-list'),
    
    path('municipio/<int:municipio_directorio_id>/', directorio_municipalidad_detail, name='directorio_municipalidad_detail'),
    
    path('municipio/public/', DirectorioMunicipioListViewPublic.as_view(), name='municipio-public'),
    
    # -- DIRECTORIO SALUD
    path('salud/create/', DirectorioSaludCreateView.as_view(), name='salud-create'),
    
    path('salud/list/', DirectorioSaludListView.as_view(), name='salud-list'),
    
    path('salud/<int:salud_directorio_id>/', directorio_salud_detail, name='directorio_salud_detail'),
    
    path('salud/public/', DirectorioSaludListViewPublic.as_view(), name='salud-public'),
    
    ####################################
    ## --- SELLO MUNICIPAL  
    ####################################    
    path('sello/', index_sello, name='index-sello'),
    ### SEGUIMIENTO
    # provincia
    path('sello_get_provincias/<int:provincias_id>/', sello_get_provincias, name='sello_get_provincias'),
    #-- provincia excel
    #path('rpt_operacional_prov_excel/', RptOperacinalProv.as_view(), name = 'rpt_operacional_prov_xls'),
    
    
    # distrito
    path('sello_get_distritos/<int:distritos_id>/', sello_get_distritos, name='sello_get_distritos'),
    path('p_distritos_sello/', sello_p_distritos, name='p_distritos_sello'),
    #-- distrito excel
    path('rpt_seguimiento_distrito_excel/', RptOperacinalDist.as_view(), name = 'rpt_seguimiento_dist_xls'),
]