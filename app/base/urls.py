from django.urls import path
from . import views 

from django.contrib import admin
from django.urls import path, include
from django.conf.urls.static import static
from django.conf import settings


urlpatterns = [
    path('', views.home_redirect_view, name='home'),
    path('inicio/', views.inicio, name='inicio'),
    path('signup/', views.signup, name='signup'),
    path('logout/', views.signout, name='logout'),
    path('signin/', views.signin, name='signin'),
    ## PADRON NOMINAL SITUACION
    path('', include('pn_situacion_actual.urls')),
    ## ACTA PADRON NOMINAL 
    path('', include('pn_acta_homologacion.urls')),
    ## NIÑOS OBSERVADOS 
    path('', include('pn_nino_observados.urls')),
    ## NIÑOS HISTORIAL 
    path('', include('pn_historial.urls')),
    ## COMPROMISO 1 NINO 
    path('', include('compromiso_nino.urls')),
    ## PACTO REGIONAL 
    path('', include('pacto.urls')),
    ## PACTO COBERTURA POBLACION 
    path('', include('pn_poblacion.urls')),
    ## PAQUETE COMPROMISO
    path('', include('paquete_compromiso.urls')), 
    ## INDICADOR COMPROMISO
    path('', include('compromiso_indicador.urls')),
]