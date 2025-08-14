from django.urls import path
from .views import index_acta_padron



urlpatterns = [
    
    path('acta_padron/', index_acta_padron, name='index_acta_padron'),
    
]