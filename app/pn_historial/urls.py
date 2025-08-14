from django.urls import path
from .views import index_historial_padron

urlpatterns = [    
    path('historial_padron/', index_historial_padron, name='index_historial_padron'),
]