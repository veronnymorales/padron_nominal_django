from django.urls import path
from .views import index_observados_padron



urlpatterns = [
    
    path('observados_padron/', index_observados_padron, name='index_observados_padron'),
    
]