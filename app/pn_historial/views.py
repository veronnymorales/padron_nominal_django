import logging
from django.shortcuts import render
from django.http import JsonResponse

logger = logging.getLogger(__name__)

def index_historial_padron(request):
    folder_id = "1Z9krFD9TLE1fqw_Wg2sMVu901tc38S3A"  # Reemplaza con tu ID real
    return render(request, "pn_historial/index_historial_padron.html", {"folder_id": folder_id})