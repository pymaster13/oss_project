from django.contrib import admin
from .models import *

@admin.register(GIP)
class GIPAdmin(admin.ModelAdmin):
    list_display = (
    'fio',
    'organization',)

    list_filter = ('fio', 'organization',)
    search_fields = ('fio', 'organization',)

@admin.register(District)
class DistrictAdmin(admin.ModelAdmin):
    list_display = (
    'name',)

    list_filter = ('name',)
    search_fields = ('name',)

@admin.register(Certificate)
class CertificateAdmin(admin.ModelAdmin):
    list_display = (
    'name',)

    list_filter = ('name',)
    search_fields = ('name',)

@admin.register(TehNadzor)
class TehNadzorAdmin(admin.ModelAdmin):
    list_display = (
    'person',)

    list_filter = ('person','district',)
    search_fields = ('person',)

@admin.register(Commission)
class CommissionAdmin(admin.ModelAdmin):
    list_display = (
    'district',)

    list_filter = ('predsedatel','predstavitel','district',)
    search_fields = ('district',)

@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = (
    'fio', 'post')

@admin.register(Object)
class ObjectAdmin(admin.ModelAdmin):
    list_display = (
    'name_object',)

admin.site.register(Order)
admin.site.register(Podzem_stal_gazoprovod)
admin.site.register(Podzem_polietilen_gazoprovod)
admin.site.register(Nadzem_stal_gazoprovod)
admin.site.register(Diametr)
admin.site.register(Diametr_x)
admin.site.register(Diametr_dlina)
admin.site.register(Diametr_kolvo)
admin.site.register(Diametr_dlina_kolvo)
admin.site.register(Diametr_x_dlina)
admin.site.register(Dop_dann_sharovyi_kran)
admin.site.register(Dop_dann_gazopr_v_zashit)
admin.site.register(Svarshik)
admin.site.register(Cokol_soed_stal_gazoprovod)
admin.site.register(Diametrs_3_dlina)
admin.site.register(Diametrs_3_kolvo)
admin.site.register(Diametr_x_dlina_prim)
admin.site.register(Neraz_soed_stal_gazoprovod)
admin.site.register(Kontragent)
admin.site.register(Smeta)
admin.site.register(Rostehnadzor)
admin.site.register(Prodavlivanie)
