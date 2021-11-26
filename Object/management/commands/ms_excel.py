import datetime

from django.core.management.base import BaseCommand
from openpyxl import load_workbook

from Object.models import *


class Command(BaseCommand):
    help = 'test MS_EXCEL'

    def handle(self, *args, **options):
        object = Object.objects.first()
        wb = load_workbook(filename = 'KS2_KS3.xlsx')

        """
        КС 2
        """

        sheet_ranges = wb['Акт по форме КС-2']
        sheet_ranges['B7'].value = sheet_ranges['B7'].value.replace( \
                    'полеЗаказчик',object.zakazchik)
        sheet_ranges['B8'].value = sheet_ranges['B8'].value.replace( \
                    'полеПодрядчик',f"{object.kontragent.name_kontragent}, \
                    ИНН {object.kontragent.INN}/ КПП {object.kontragent.KPP}, \
                     адрес: {object.kontragent.Ur_address}, тел:{object.kontragent.telephone}")
        sheet_ranges['C9'].value = sheet_ranges['C9'].value.replace( \
                    'полеСтройка',object.ks2_stroika)
        sheet_ranges['C10'].value = sheet_ranges['C10'].value.replace( \
                    'полеОбъект',object.ks2_object)
        sheet_ranges['F12'].value = sheet_ranges['F12'].value.replace( \
                    'полеНомердоговоранаразм',object.Nomer_razm)
        sheet_ranges['F13'].value = sheet_ranges['F13'].value.replace( \
                    'полеДатадоговоранаразм',datetime.datetime.strftime(object.Data_razm,"%d.%m.%Y"))
        sheet_ranges['F18'].value = sheet_ranges['F18'].value.replace( \
                    'полеДатазам1',datetime.datetime.strftime(object.Data_zamera1,"%d.%m.%Y"))
        sheet_ranges['G18'].value = sheet_ranges['G18'].value.replace( \
                    'полеДатазам1',datetime.datetime.strftime(object.Data_zamera1,"%d.%m.%Y"))
        sheet_ranges['H18'].value = sheet_ranges['H18'].value.replace( \
                    'полеДатазам2',datetime.datetime.strftime(object.Data_zamera2,"%d.%m.%Y"))

        count = 1
        start_value = ''
        start_position = ''
        
        while (start_value == ''):
            if sheet_ranges[f'A{count}'].value:
                if 'Сдал:' in str(sheet_ranges[f'A{count}'].value):
                    start_value = 'Сдал:'
                    start_position = f'A{count}'
                    sheet_ranges[f'A{count}'].value = sheet_ranges[f'A{count}'].value.replace( \
                                'полеКС2подрядчик',object.ks2_podryadchik.post)
                    sheet_ranges[f'F{count}'].value = sheet_ranges[f'F{count}'].value.replace( \
                                'полеКС2подрядчикФИО',object.ks2_podryadchik.fio)
                else:
                    count += 1
            else:
                count += 1

        while (start_value == 'Сдал:'):
            if sheet_ranges[f'A{count}'].value:
                if 'Принял:' in str(sheet_ranges[f'A{count}'].value):
                    start_value = 'Принял:'
                    start_position = f'A{count}'
                    sheet_ranges[f'A{count}'].value = sheet_ranges[f'A{count}'].value.replace( \
                                'полеКС2заказчик',object.ks2_zakazchik.post)
                    sheet_ranges[f'F{count}'].value = sheet_ranges[f'F{count}'].value.replace( \
                                'полеКС2заказчикФИО',object.ks2_zakazchik.fio)
                else:
                    count += 1
            else:
                count += 1

        while (start_value == 'Принял:'):
            if sheet_ranges[f'A{count}'].value:
                if 'полеТехнадзор' in str(sheet_ranges[f'A{count}'].value):
                    start_value = 'полеТехнадзор'
                    start_position = f'A{count}'
                    sheet_ranges[f'A{count}'].value = sheet_ranges[f'A{count}'].value.replace( \
                                'полеТехнадзор',object.tehnadzor.person.post)
                    sheet_ranges[f'F{count}'].value = sheet_ranges[f'F{count}'].value.replace( \
                                'полеТехнадзорФИО',object.tehnadzor.person.fio)
                else:
                    count += 1
            else:
                count += 1


        """
        КС 3
        """

        sheet_ranges_ks3 = wb['КС3']
        sheet_ranges_ks3['A9'].value = sheet_ranges_ks3['A9'].value.replace( \
                    'полеЗаказчик',object.zakazchik)
        sheet_ranges_ks3['A11'].value = sheet_ranges_ks3['A11'].value.replace( \
                    'полеПодрядчик',f"{object.kontragent.name_kontragent}, \
                    ИНН {object.kontragent.INN}/ КПП {object.kontragent.KPP}, \
                    адрес: {object.kontragent.Ur_address}, тел:{object.kontragent.telephone}")

        sheet_ranges_ks3['A13'].value = sheet_ranges_ks3['A13'].value.replace( \
                                                    'полеСтройка',object.ks2_stroika)
        sheet_ranges_ks3['A15'].value = sheet_ranges_ks3['A15'].value.replace( \
                                                    'полеОбъект',object.ks2_object)

        sheet_ranges_ks3['I17'].value = sheet_ranges_ks3['I17'].value.replace( \
                    'полеНомердоговоранаразм',object.Nomer_razm)
        sheet_ranges_ks3['I18'].value = sheet_ranges_ks3['I18'].value.replace( \
                    'дд',datetime.datetime.strftime(object.Data_razm,"%d"))
        sheet_ranges_ks3['J18'].value = sheet_ranges_ks3['J18'].value.replace( \
                    'мм',datetime.datetime.strftime(object.Data_razm,"%m"))
        sheet_ranges_ks3['K18'].value = sheet_ranges_ks3['K18'].value.replace( \
                    'гггг',datetime.datetime.strftime(object.Data_razm,"%Y"))

        sheet_ranges_ks3['A40'].value = sheet_ranges_ks3['A40'].value.replace( \
                    'полеКС2подрядчик',object.ks2_podryadchik.post)
        sheet_ranges_ks3['H40'].value = sheet_ranges_ks3['H40'].value.replace( \
                    'полеКС2подрядчикФИО',object.ks2_podryadchik.fio)
        sheet_ranges_ks3['A46'].value = sheet_ranges_ks3['A46'].value.replace( \
                    'полеКС2заказчик',object.ks2_zakazchik.post)
        sheet_ranges_ks3['H46'].value = sheet_ranges_ks3['H46'].value.replace( \
                    'полеКС2заказчикФИО',object.ks2_zakazchik.fio)

        sheet_ranges_ks3['F23'].value = sheet_ranges_ks3['F23'].value.replace( \
                    'полеДатазам1',datetime.datetime.strftime(object.Data_zamera1,"%d.%m.%Y"))
        sheet_ranges_ks3['H23'].value = sheet_ranges_ks3['H23'].value.replace( \
                    'полеДатазам1',datetime.datetime.strftime(object.Data_zamera1,"%d.%m.%Y"))
        sheet_ranges_ks3['I23'].value = sheet_ranges_ks3['I23'].value.replace( \
                    'полеДатазам2',datetime.datetime.strftime(object.Data_zamera2,"%d.%m.%Y"))
        sheet_ranges_ks3['B31'].value = sheet_ranges_ks3['B31'].value.replace( \
                                                    'полеОбъект',object.name_object)


        wb.save('КС2-КС3 - {}.xlsx'.format(object.name_object))
