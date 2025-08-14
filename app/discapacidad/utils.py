from base.models import DimPeriodo, DimDiscapacidadEtapa, MAESTRO_HIS_ESTABLECIMIENTO
from .models import DimDisFisicaCie,TramaBaseDiscapacidadRpt02FisicaNominal
from django.db.models import Case, When, Value, IntegerField, Sum

def generar_operacional():
    rpt_operacional = TramaBaseDiscapacidadRpt02FisicaNominal.objects.values('renaes') \
            .annotate(
                dis_1=Sum(
                    Case(
                        When(Categoria=1, gedad=1, then=Value(1)),
                        default=Value(0),
                        output_field=IntegerField()
                    )
                ),
                dis_2=Sum(
                    Case(
                        When(Categoria=1, gedad=2, then=Value(1)),
                        default=Value(0),
                        output_field=IntegerField()
                    )
                ),
                dis_3=Sum(
                    Case(
                        When(Categoria=1, gedad=3, then=Value(1)),
                        default=Value(0),
                        output_field=IntegerField()
                    )
                ),
                dis_4=Sum(
                    Case(
                        When(Categoria=1, gedad=4, then=Value(1)),
                        default=Value(0),
                        output_field=IntegerField()
                    )
                ),
                dis_5=Sum(
                    Case(
                        When(Categoria=1, gedad=5, then=Value(1)),
                        default=Value(0),
                        output_field=IntegerField()
                    )
                )
            )
    return rpt_operacional

