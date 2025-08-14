from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib.auth.models import User
from django.contrib.auth import login, logout, authenticate
from django.shortcuts import render, redirect
from django.db import connection

#POI
from django.db import transaction
from .poi_models import ActividadPOI, ProgramacionMensual, ProgramacionMetaFisica
from .forms import ActividadPOIForm, ProgramacionMensualFormSet
from django.core.paginator import Paginator
from django.db.models import Q
from .forms import RegistroTareaForm, ProgramacionMensualMetaFisicaForm, ProgramacionMensualMetaFinancieraForm


############################################
########### POI ############################
############################################
@transaction.atomic
def registro_actividad_poi(request):
    if request.method == 'POST':
        form = ActividadPOIForm(request.POST)
        if form.is_valid():
            actividad_poi = form.save()
            formset = ProgramacionMensualFormSet(request.POST, queryset=ProgramacionMensual.objects.none())
            if formset.is_valid():
                instances = formset.save(commit=False)
                for instance in instances:
                    instance.actividad_poi = actividad_poi
                    instance.save()
                return redirect('lista_actividades_poi')
    else:
        form = ActividadPOIForm()
        formset = ProgramacionMensualFormSet(queryset=ProgramacionMensual.objects.none())

    return render(request, 'discapacidad/registro_actividad_poi.html', {'form': form, 'formset': formset})

def lista_actividades_poi(request):
    actividades_list = ActividadPOI.objects.all()
    if request.GET.get('q'):
        actividades_list = actividades_list.filter(Q(actividad_presupuestal__icontains=request.GET['q']) | Q(producto_presupuestal__icontains=request.GET['q']) | Q(unidad_ejecutora__icontains=request.GET['q']))
    if request.GET.get('ano'):
        actividades_list = actividades_list.filter(ano=request.GET['ano'])
    if request.GET.get('tipo_categoria'):
        actividades_list = actividades_list.filter(tipo_categoria=request.GET['tipo_categoria'])

    paginator = Paginator(actividades_list, 20)  # 20 actividades por página
    page = request.GET.get('page')
    actividades = paginator.get_page(page)

    context = {
        'actividades': actividades,
        'anos_disponibles': ActividadPOI.objects.values_list('ano', flat=True).distinct(),
        'tipos_categoria': ActividadPOI.TIPO_CATEGORIA,
    }
    return render(request, 'discapacidad/lista_actividades_poi.html', context)

def registro_actividad_detail(request,registro_actividad_id):
    if request.method == 'GET':
        registro_poi = get_object_or_404(ActividadPOI, pk=registro_actividad_id)
        form = ActividadPOIForm(instance=registro_poi)
        context = {
            'registro_poi': registro_poi,
            'form': form
        }
        return render(request, 'discapacidad/registro_actividad_detail.html', context)
    else:
        try:
            registro_poi = get_object_or_404(ActividadPOI, pk=registro_actividad_id)
            form = ActividadPOIForm(request.POST, instance=registro_poi)
            form.save()
            return redirect('lista_actividades_poi')
        except ValueError:
            return render(request, 'discapacidad/registro_actividad_detail.html', {'registro_poi': registro_poi, 'form': form, 'error': 'Error actualizar'})


def registrar_tarea(request,registro_actividad_id):
    if request.method == 'POST':
        tarea_form = RegistroTareaForm(request.POST)
        fisica_form = ProgramacionMensualMetaFisicaForm(request.POST)
        financiera_form = ProgramacionMensualMetaFinancieraForm(request.POST)

        if tarea_form.is_valid() and fisica_form.is_valid() and financiera_form.is_valid():
            tarea = tarea_form.save()
            fisica = fisica_form.save(commit=False)
            financiera = financiera_form.save(commit=False)

            fisica.registro_tarea = tarea
            financiera.registro_tarea = tarea

            fisica.total = (fisica.enero + fisica.febrero + fisica.marzo + fisica.abril + 
                            fisica.mayo + fisica.junio + fisica.julio + fisica.agosto + 
                            fisica.setiembre + fisica.octubre + fisica.noviembre + fisica.diciembre)

            financiera.total = (financiera.enero + financiera.febrero + financiera.marzo + financiera.abril + 
                                financiera.mayo + financiera.junio + financiera.julio + financiera.agosto + 
                                financiera.setiembre + financiera.octubre + financiera.noviembre + financiera.diciembre)

            fisica.save()
            financiera.save()

            return redirect('discapacidad/registro_actividad_detail.html', registro_actividad_detail_id=tarea.id)
    else:
        tarea_form = RegistroTareaForm()
        fisica_form = ProgramacionMensualMetaFisicaForm()
        financiera_form = ProgramacionMensualMetaFinancieraForm()

    return render(request, 'registro_tarea.html', {
        'tarea_form': tarea_form,
        'fisica_form': fisica_form,
        'financiera_form': financiera_form,
    })