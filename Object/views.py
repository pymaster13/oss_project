import json
import datetime

from django.shortcuts import render
import docx

from .models import *
from .ms_word_ITD import form_ITD
from .ms_word_ks14 import form_ks14
from .ms_excel_ks2 import form_ks2
from .ms_word_dop_rostehnadzor import form_dop_ITD

fields = {'name_object':'Название объекта', 'zakazchik':'Заказчик', 'place': 'Местоположение',
       'kod_object':'Код объекта', 'district':'Район', 'kontragent':'Контрагент',
       'Nomer_proekt':'Номер проекта', 'Nomer_razm':'Номер договора на размещение',
       'Data_razm':'Дата договора на размещение',
       'Nomer_zadaniya':'Номер задания проектирования', 'Date_proekt':'Дата проектирования',
       'gip':'ГИП (ФИО, организация)', 'zayv':'Заявитель',
       'ks2':'Подписанты КС 2 и КС 3 (+ объект, стройка)', 'ks11':'Подписанты КС 11 и КС 14',
       'prover_davl':'Проверочное давление', 'davl':'Давление',
       'gazoprovod_podzem_stal':'Подземный стальной газопровод',
       'gazoprovod_podzem_poliet':'Подземный полиэтиленовый газопровод',
       'gazoprovod_nadzem_stal':'Надземный стальной газопровод',
       'certificates':'Сертификаты', 'gazoprovod':'Газопровод (ЭХЗ)',
       'prodavlivanie':'Продавливание', 'svarshik1':'Сварщик (полиэтилен)',
       'svarshik2':'Сварщик (сталь)', 'tehnadzor': 'Технадзор',
       'proektnaya_org':'Проектная организация', 'Data_zamera1':'Дата замера 1',
       'Data_zamera2':'Дата замера 2', 'Data_sost_project':'Дата составления проекта',
       'Data_razbiv':'Дата разбивки', 'Data_produv':'Дата продувки',
       'Data_ukl':'Дата укладки', 'zashitnyi_futlyar':'Защитный футляр',
       'futlyar_na_vyhode':'Футляр на выходе земли', 'opora':'Опора под газопровод',
       'dop_dann_sharovyi_kran':'Доп. данные (шаровый кран)',
       'dop_dann_gazopr_v_zashit':'Доп. данные (прокладка газопровода)',
       'dop_dann_futlyar_na_vyhode':'Доп. данные (футляр на выходе из земли)',
       'dop_dann_ob_ustanovke_opor':'Доп. данные (установка опор)',
       'smeta_data':'Информация от сметчика', 'rostehnadzor':'Доп. данные (ростехнадзор)',
       'prodavl':'Доп. данные (продавливание)'
       }

fields_dates = {'Data_razm':'Дата договора на размещение', 'Date_proekt':'Дата проектирования',
       'Data_zamera1':'Дата замера 1', 'Data_zamera2':'Дата замера 2',
       'Data_sost_project':'Дата составления проекта',
       'Data_razbiv':'Дата разбивки', 'Data_produv':'Дата продувки',
       'Data_ukl':'Дата укладки'
       }


def index(request):
    return render(request, 'index.html', {})


def object_add(request):
    context = {}
    gips_orgs = {}
    gips_list = []
    raions_list = []

    fill_all_lists(context, gips_orgs, gips_list, raions_list)

    if request.method == "GET":
        return render(request, 'object_add.html', {'context':context})

    else:
        data = request.POST
        if not data['name_object']:
            context['error'] = 'Название объекта не может быть пустым!'
            return render(request, 'object_add.html', {'context':context})

        keys = data.keys()

        indexes_podzem_stalnoi_truba = []
        indexes_poliet_truba_mufta = []
        indexes_nadzem_stalnoi_truba = []
        indexes_dop_dann_shar_kran = []
        indexes_dop_dann_gazopr_v_zashit = []
        indexes_dop_dann_ob_ustanovke_futl = []
        indexes_dop_dann_opor_pod = []

        for key in keys:
            if "stalnoi_truba_diametr" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_podzem_stalnoi_truba.append(splitted[1])
            if "poliet_truba_dlina" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_poliet_truba_mufta.append(splitted[1])
            if "nadzem_stal_truba_diametr" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_nadzem_stalnoi_truba.append(splitted[1])
            if "dop_dann_sharovyi_kran_mufta" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_dop_dann_shar_kran.append(splitted[1])
            if "dop_dann_gazopr_v_zashit_diametr" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_dop_dann_gazopr_v_zashit.append(splitted[1])
            if "dop_dann_ob_ustanovke_futlyarov_truba_diametr" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_dop_dann_ob_ustanovke_futl.append(splitted[1])
            if "dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr" in key:
                splitted = key.split('-')
                if len(splitted) > 1:
                    indexes_dop_dann_opor_pod.append(splitted[1])

        object_ = Object.objects.get_or_create(name_object=data['name_object'])
        if not object_[1]:
            context['error'] = 'Проверьте корректность названия объекта!'
            return render(request, 'object_add.html', {'context':context})

        object = object_[0]

        if data['name_object']:
            object.name_object = data['name_object']
        if data['zakazchik']:
            object.zakazchik = data['zakazchik']
        if data['place']:
            object.place = data['place']
        if data['kod_object']:
            object.kod_object = data['kod_object']

        if data['select_raion'] != "Выберите район":
            district = District.objects.get(name = data['select_raion'])
            object.district = district
        else:
            district = District.objects.get_or_create(name = data['select_raion_custom'])[0]
            district.save()
            object.district = district

        if data['name_kontragent'] or data['INN'] or data['KPP'] or data['telephone'] \
            or data['Ur_address'] or data['ks11_predstav_podryadchika_fio'] \
            or data['ks11_predstav_podryadchika_post']:

            kontragent = Kontragent.objects.get_or_create(
                name_kontragent = data['name_kontragent'])[0]
            kontragent.INN = data['INN']
            kontragent.KPP = data['KPP']
            kontragent.telephone = data['telephone']
            kontragent.Ur_address = data['Ur_address']
            podpisant = Person.objects.get_or_create(fio=data['ks11_predstav_podryadchika_fio'],
                                                post=data['ks11_predstav_podryadchika_post'])[0]
            podpisant.save()
            kontragent.podpisant = podpisant
            kontragent.save()
            object.kontragent = kontragent

        if data['zayv']:
            object.zayv = data['zayv']

        if data['Data_razm']:
            object.Data_razm = datetime.datetime.strptime(data['Data_razm'],"%Y-%m-%d")

        if data['Nomer_razm']:
            object.Nomer_razm = data['Nomer_razm']
        if data['Nomer_zadaniya']:
            object.Nomer_zadaniya = data['Nomer_zadaniya']
        if data['Nomer_proekt']:
            object.Nomer_proekt = data['Nomer_proekt']

        if data['Date_proekt']:
            object.Date_proekt = datetime.datetime.strptime(data['Date_proekt'],"%Y-%m-%d")

        if data['select_GIP'] != "Выберите ГИПа":
            gip = GIP.objects.get(fio = data['select_GIP'])

        else:
            gip = GIP.objects.get_or_create(fio = data['custom_GIP'])[0]

        gip.organization = data['organization']
        gip.save()
        object.gip = gip

        if data['ks2_zakazchik_fio'] or data['ks2_zakazchik_post']:
            person1 = Person.objects.get_or_create(fio=data['ks2_zakazchik_fio'],
                                                post=data['ks2_zakazchik_post'])[0]
            person1.save()
            object.ks2_zakazchik = person1

        if data['ks2_podryadchik_fio'] or data['ks2_podryadchik_post']:
            person2 = Person.objects.get_or_create(fio=data['ks2_podryadchik_fio'],
                                                post=data['ks2_podryadchik_post'])[0]
            person2.save()
            object.ks2_podryadchik = person2

        if data['ks11_predsedatel_fio'] or data['ks11_predsedatel_post']:
            person3 = Person.objects.get_or_create(fio=data['ks11_predsedatel_fio'],
                                                post=data['ks11_predsedatel_post'])[0]
            person3.save()
            object.ks11_predsedatel = person3

        if data['object']:
            object.ks2_object = data['object']

        if data['stroika']:
            object.ks2_stroika = data['stroika']

        if data['ks11_predstav_proekt_fio'] or data['ks11_predstav_proekt_post']:
            person4 = Person.objects.get_or_create(fio=data['ks11_predstav_proekt_fio'],
                                                post=data['ks11_predstav_proekt_post'])[0]
            person4.save()
            object.ks11_predstav_proekt = person4

        if data['ks11_predstav_ekspl_fio'] or data['ks11_predstav_ekspl_post']:
            person5 = Person.objects.get_or_create(fio=data['ks11_predstav_ekspl_fio'],
                                                post=data['ks11_predstav_ekspl_post'])[0]
            person5.save()
            object.ks11_predstav_ekspl = person5

        if data['select_davlenie'] != 'Выберите давление, МПа':
            object.prover_davl = data['select_davlenie']
            if object.prover_davl == '0.3':
                object.davlenie_name = "низкого давления"
            if object.prover_davl == '0.6':
                object.davlenie_name = "среднего давления"
        else:
            object.prover_davl = data['select_davlenie_custom']
            object.davlenie_name = data['davlenie_name']

        if data['davl']:
            object.davl = data['davl']

        """   Подземный стальной газопровод   """

        podzem_stal_gazoprovod = Podzem_stal_gazoprovod()

        existed_new_neraz_soed_truba_podzem_stal_gazoprovod = ''

        if not Neraz_soed_stal_gazoprovod.objects.get_or_create(
            PE = data['stalnoi_neraz_soed_PE'],
            ST = data['stalnoi_neraz_soed_ST'],
            kolvo = data['neraz_soed_kolvo'])[1]:

            existed_new_neraz_soed_truba_podzem_stal_gazoprovod = Neraz_soed_stal_gazoprovod.objects.get(
                PE = data['stalnoi_neraz_soed_PE'],
                ST = data['stalnoi_neraz_soed_ST'],
                kolvo = data['neraz_soed_kolvo'])

        if existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
            podzem_stal_gazoprovod.neraz_soed = existed_new_neraz_soed_truba_podzem_stal_gazoprovod
        else:
            neraz_soed =  Neraz_soed_stal_gazoprovod.objects.get_or_create(
                PE = data['stalnoi_neraz_soed_PE'],
                ST = data['stalnoi_neraz_soed_ST'],
                kolvo = data['neraz_soed_kolvo'])[0]
            neraz_soed.save()
            podzem_stal_gazoprovod.neraz_soed = neraz_soed

        podzem_stal_gazoprovod.kontrolnaya_trubka = data['stalnoi_kontrolnaya_trubka']
        podzem_stal_gazoprovod.otvod_90 = data['stalnoi_otvod']
        stal_opoznavat_znak = data['stal_opoznavat_znak']
        podzem_stal_gazoprovod.opoznavat_znak = stal_opoznavat_znak

        podzem_stal_gazoprovod.save()
        if 'stalnoi_truba_diametr' in keys:
            _add_truba_in_podzem_stal_gazoprovod(data=data,object=podzem_stal_gazoprovod)
        if indexes_podzem_stalnoi_truba:
            for index in indexes_podzem_stalnoi_truba:
                _add_truba_in_podzem_stal_gazoprovod(data=data,object=podzem_stal_gazoprovod,
                                                        index='-{}'.format(index))

        a = Podzem_stal_gazoprovod.objects.filter(
            neraz_soed=podzem_stal_gazoprovod.neraz_soed,
            kontrolnaya_trubka=data['stalnoi_kontrolnaya_trubka'],
            otvod_90=data['stalnoi_otvod'],
            opoznavat_znak=stal_opoznavat_znak)

        if a:
            if len(a) > 1:
                for elem in a:
                    if (elem.pk != podzem_stal_gazoprovod.pk and list(elem.truba.all()) == list(podzem_stal_gazoprovod.truba.all())):
                        object.gazoprovod_podzem_stal = elem
                        podzem_stal_gazoprovod.delete()
                        break

            else:
                podzem_stal_gazoprovod.save()
                object.gazoprovod_podzem_stal = podzem_stal_gazoprovod

        else:
            podzem_stal_gazoprovod.save()
            object.gazoprovod_podzem_stal = podzem_stal_gazoprovod

        """   Подземный полиэтиленовый газопровод   """

        podzem_polietilen_gazoprovod = Podzem_polietilen_gazoprovod()

        existed_otvod = ''

        if not Diametr_kolvo.objects.get_or_create(diametr= data['poliet_otvod_diametr'],
                                                    kolvo=data['poliet_otvod_dlina'])[1]:
            existed_otvod = Diametr_kolvo.objects.get(diametr= data['poliet_otvod_diametr'],
                                                        kolvo=data['poliet_otvod_dlina'])

        if existed_otvod:
            podzem_polietilen_gazoprovod.otvod = existed_otvod
        else:
            otvod = Diametr_kolvo.objects.get_or_create(diametr= data['poliet_otvod_diametr'],
                                                        kolvo=data['poliet_otvod_dlina'])[0]
            otvod.save()
            podzem_stal_gazoprovod.otvod = otvod

        existed_troinik = ''

        if not Diametrs_3_kolvo.objects.get_or_create(diametr1= data['poliet_troinik1'], \
                        diametr2= data['poliet_troinik2'], diametr3= data['poliet_troinik3'],
                        kolvo=data['poliet_troinik_dlina'])[1]:
            existed_troinik = Diametrs_3_kolvo.objects.get(diametr1= data['poliet_troinik1'], \
                            diametr2= data['poliet_troinik2'], diametr3= data['poliet_troinik3'],
                            kolvo=data['poliet_troinik_dlina'])

        if existed_troinik:
            podzem_polietilen_gazoprovod.troinik = existed_troinik
        else:
            troinik = Diametrs_3_kolvo.objects.get_or_create(diametr1= data['poliet_troinik1'], \
                            diametr2= data['poliet_troinik2'], diametr3= data['poliet_troinik3'],
                            kolvo=data['poliet_troinik_dlina'])[0]
            troinik.save()
            podzem_polietilen_gazoprovod.troinik = troinik

        existed_sedelka = ''

        new_sedelka = Diametrs_3_kolvo.objects.get_or_create(diametr1= data['sedelka_poliet_troinik1'], \
            diametr2= data['sedelka_poliet_troinik2'], diametr3= data['sedelka_poliet_troinik3'],
            kolvo = data['sedelka_poliet_kolvo'])

        if not new_sedelka[1]:
            existed_sedelka = new_sedelka[0]

        if existed_sedelka:
            podzem_polietilen_gazoprovod.sedelka = existed_sedelka
        else:
            sedelka = Diametrs_3_kolvo.objects.get_or_create(diametr1= data['sedelka_poliet_troinik1'], \
                diametr2= data['sedelka_poliet_troinik2'], diametr3= data['sedelka_poliet_troinik3'],
                kolvo = data['sedelka_poliet_kolvo'])[0]
            sedelka.save()
            podzem_polietilen_gazoprovod.sedelka = sedelka

        existed_zaglushka = ''

        new_zaglushka = Diametr_kolvo.objects.get_or_create(diametr= data['poliet_zaglushka_diametr'],
                                                        kolvo=data['poliet_zaglushka_dlina'])

        if not new_zaglushka[1]:
            existed_zaglushka = new_zaglushka[0]

        if existed_zaglushka:
            podzem_polietilen_gazoprovod.zaglushka = existed_zaglushka
        else:
            zaglushka = Diametr_kolvo.objects.get_or_create(diametr= data['poliet_zaglushka_diametr'],
                                                            kolvo=data['poliet_zaglushka_dlina'])[0]
            zaglushka.save()
            podzem_polietilen_gazoprovod.zaglushka = zaglushka

        existed_kran = ''

        new_kran = Diametr_kolvo.objects.get_or_create(diametr= data['poliet_kran_shar_diametr'],
                                                    kolvo=data['poliet_kran_shar_dlina'])

        if not new_kran[1]:
            existed_kran = new_kran[0]

        if existed_kran:
            podzem_polietilen_gazoprovod.kran = existed_kran
        else:
            kran = Diametr_kolvo.objects.get_or_create(diametr= data['poliet_kran_shar_diametr'],
                                                        kolvo=data['poliet_kran_shar_dlina'])[0]
            kran.save()
            podzem_polietilen_gazoprovod.kran = kran

        podzem_polietilen_gazoprovod.lenta = data['poliet_lenta_signal_dlina']
        podzem_polietilen_gazoprovod.znak = data['poliet_opoznavat_znak_kolvo']

        podzem_polietilen_gazoprovod.save()

        if 'poliet_truba_dlina' in keys:
            _add_poliet_truba_mufta_in_podzem_poliet_gazoprovod(data=data,
                                            object=podzem_polietilen_gazoprovod)

        if indexes_poliet_truba_mufta:
            for index in indexes_poliet_truba_mufta:
                _add_poliet_truba_mufta_in_podzem_poliet_gazoprovod(data=data,
                                            object=podzem_polietilen_gazoprovod,
                                            index='-{}'.format(index))

        b = Podzem_polietilen_gazoprovod.objects.filter(
            otvod=podzem_polietilen_gazoprovod.otvod,
            troinik=podzem_polietilen_gazoprovod.troinik,
            sedelka=podzem_polietilen_gazoprovod.sedelka,
            zaglushka=podzem_polietilen_gazoprovod.zaglushka,
            kran=podzem_polietilen_gazoprovod.kran,
            lenta=podzem_polietilen_gazoprovod.lenta,
            znak=podzem_polietilen_gazoprovod.znak)

        if b:
            if len(b) > 1:
                for elem in b:
                    if (elem.pk != podzem_polietilen_gazoprovod.pk
                        and list(elem.truba.all()) == list(podzem_polietilen_gazoprovod.truba.all())
                        and list(elem.mufta.all()) == list(podzem_polietilen_gazoprovod.mufta.all())):

                        object.gazoprovod_podzem_poliet = elem
                        podzem_polietilen_gazoprovod.delete()
                        break

            else:
                podzem_polietilen_gazoprovod.save()
                object.gazoprovod_podzem_poliet = podzem_polietilen_gazoprovod

        else:
            podzem_polietilen_gazoprovod.save()
            object.gazoprovod_podzem_poliet = podzem_polietilen_gazoprovod


        """  Надземный_стальной_газопровод  """

        nadzem_stal_gazoprovod = Nadzem_stal_gazoprovod()

        existed_nadzem_stal_gazoprovod_izolir = ''

        new_nadzem_stal_gazoprovod_izolir = Diametr_kolvo.objects.get_or_create(\
                                                diametr = data['nadzem_stal_izolir_soed'],\
                                                kolvo = data['nadzem_stal_izolir_kolvo'])

        if not new_nadzem_stal_gazoprovod_izolir[1]:
            existed_nadzem_stal_gazoprovod_izolir = new_nadzem_stal_gazoprovod_izolir[0]

        if existed_nadzem_stal_gazoprovod_izolir:
            nadzem_stal_gazoprovod.izolir_soed = existed_nadzem_stal_gazoprovod_izolir

        else:
            nadzem_stal_gazoprovod_izolir = Diametr_kolvo.objects.get_or_create(\
                                                    diametr = data['nadzem_stal_izolir_soed'],\
                                                    kolvo = data['nadzem_stal_izolir_kolvo'])[0]
            nadzem_stal_gazoprovod_izolir.save()

            nadzem_stal_gazoprovod.izolir_soed = nadzem_stal_gazoprovod_izolir

        existed_nadzem_stal_gazoprovod_kran = ''

        new_nadzem_stal_gazoprovod_kran = Diametr_kolvo.objects.get_or_create( \
                                                diametr = data['nadzem_stal_kran_diametr'],\
                                                kolvo = data['nadzem_stal_kran_kolvo'])

        if not new_nadzem_stal_gazoprovod_kran[1]:
            existed_nadzem_stal_gazoprovod_kran = new_nadzem_stal_gazoprovod_kran[0]

        if existed_nadzem_stal_gazoprovod_kran:
            nadzem_stal_gazoprovod.kran_stal = existed_nadzem_stal_gazoprovod_kran
        else:
            nadzem_stal_gazoprovod_kran = Diametr_kolvo.objects.get_or_create( \
                                                    diametr = data['nadzem_stal_kran_diametr'],\
                                                    kolvo = data['nadzem_stal_kran_kolvo'])[0]
            nadzem_stal_gazoprovod_kran.save()

            nadzem_stal_gazoprovod.kran_stal = nadzem_stal_gazoprovod_kran


        existed_nadzem_stal_gazoprovod_otvod = ''

        new_nadzem_stal_gazoprovod_otvod = Diametr_kolvo.objects.get_or_create(\
                                                diametr = data['nadzem_stal_otvod_diametr'],\
                                                kolvo = data['nadzem_stal_otvod_kolvo'])

        if not new_nadzem_stal_gazoprovod_otvod[1]:
            existed_nadzem_stal_gazoprovod_otvod = new_nadzem_stal_gazoprovod_otvod[0]

        if existed_nadzem_stal_gazoprovod_otvod:
            nadzem_stal_gazoprovod.otvod = existed_nadzem_stal_gazoprovod_otvod
        else:
            nadzem_stal_gazoprovod_otvod = Diametr_kolvo.objects.get_or_create(\
                                                    diametr = data['nadzem_stal_otvod_diametr'],\
                                                    kolvo = data['nadzem_stal_otvod_kolvo'])[0]
            nadzem_stal_gazoprovod_otvod.save()

            nadzem_stal_gazoprovod.otvod = nadzem_stal_gazoprovod_otvod

        existed_nadzem_stal_gazoprovod_cokol = ''

        new_nadzem_stal_gazoprovod_cokol = Cokol_soed_stal_gazoprovod.objects.get_or_create(\
                                                PE = data['nadzem_stal_cokol_vvod_PE'],\
                                                ST = data['nadzem_stal_cokol_vvod_ST'], \
                                                kolvo = data['nadzem_stal_cokol_vvod_kolvo'])

        if not new_nadzem_stal_gazoprovod_cokol[1]:
            existed_nadzem_stal_gazoprovod_cokol = new_nadzem_stal_gazoprovod_cokol[0]

        if existed_nadzem_stal_gazoprovod_cokol:
            nadzem_stal_gazoprovod.cokolnyi_vvod = existed_nadzem_stal_gazoprovod_cokol
        else:
            nadzem_stal_gazoprovod_cokol = Cokol_soed_stal_gazoprovod.objects.get_or_create(\
                                                    PE = data['nadzem_stal_cokol_vvod_PE'],\
                                                    ST = data['nadzem_stal_cokol_vvod_ST'], \
                                                    kolvo = data['nadzem_stal_cokol_vvod_kolvo'])[0]
            nadzem_stal_gazoprovod_cokol.save()

            nadzem_stal_gazoprovod.cokolnyi_vvod = nadzem_stal_gazoprovod_cokol

        nadzem_stal_gazoprovod.kreplenie = data['nadzem_stal_kreplenie']

        existed_stoika = ''

        new_stoika = Diametr_dlina_kolvo.objects.get_or_create(diametr=data['nadzem_stalnoi_stoika_diametr'],
                                dlina=data['nadzem_stalnoi_stoika_dlina'], kolvo=data['stoika'])

        if not new_stoika[1]:
            existed_stoika = new_stoika[0]

        if existed_stoika:
            nadzem_stal_gazoprovod.stoika = existed_stoika
        else:
            stoika = Diametr_dlina_kolvo.objects.get_or_create(diametr=data['nadzem_stalnoi_stoika_diametr'],
                                    dlina=data['nadzem_stalnoi_stoika_dlina'], kolvo=data['stoika'])[0]
            stoika.save()
            nadzem_stal_gazoprovod.stoika = stoika

        nadzem_stal_gazoprovod.save()

        if 'nadzem_stal_truba_diametr' in keys:
            _add_truba_in_nadzem_stal_gazoprovod(data=data,object=nadzem_stal_gazoprovod)

        if indexes_nadzem_stalnoi_truba:
            for index in indexes_nadzem_stalnoi_truba:
                _add_truba_in_nadzem_stal_gazoprovod(data=data,object=nadzem_stal_gazoprovod,
                                                    index='-{}'.format(index))

        c = Nadzem_stal_gazoprovod.objects.filter(
            izolir_soed=nadzem_stal_gazoprovod.izolir_soed,
            kran_stal=nadzem_stal_gazoprovod.kran_stal,
            otvod=nadzem_stal_gazoprovod.otvod,
            cokolnyi_vvod=nadzem_stal_gazoprovod.cokolnyi_vvod,
            kreplenie=nadzem_stal_gazoprovod.kreplenie,
            stoika=nadzem_stal_gazoprovod.stoika)

        if c:
            if len(c) > 1:
                for elem in c:
                    if (elem.pk != nadzem_stal_gazoprovod.pk
                        and list(elem.truba.all()) == list(nadzem_stal_gazoprovod.truba.all())):
                        object.gazoprovod_nadzem_stal = elem
                        nadzem_stal_gazoprovod.delete()
                        break

            else:
                nadzem_stal_gazoprovod.save()
                object.gazoprovod_nadzem_stal = nadzem_stal_gazoprovod

        else:
            nadzem_stal_gazoprovod.save()
            object.gazoprovod_nadzem_stal = nadzem_stal_gazoprovod

        object.save()

        if 'select_cert' in keys:
            for cert in data.getlist('select_cert'):
                certificate = Certificate.objects.get(name=cert)
                certificate.save()
                object.certificates.add(certificate)

        if 'gazoprovod_input' in keys:
            object.gazoprovod = True

        if 'prodavlivanie_input' in keys:
            object.prodavlivanie = True

        if data['date_polietilens']:
            date_polietilens = datetime.datetime.strptime(data['date_polietilens'],"%Y-%m-%d")
            svarshik1 = Svarshik.objects.get_or_create(fio=data['svarshik_polietilens'],\
                            type='Полиэтилен', diametr=data['diametr_poliet_styki'],
                            kolvo=data['kolvo_poliet_styki'],
                            date_svarki=date_polietilens)[0]
        else:
            svarshik1 = Svarshik.objects.get_or_create(fio=data['svarshik_polietilens'],\
                            type='Полиэтилен', diametr=data['diametr_poliet_styki'], \
                            kolvo=data['kolvo_poliet_styki'])[0]

        svarshik1.save()
        object.svarshik1 = svarshik1

        if data['date_stal']:
            date_stal = datetime.datetime.strptime(data['date_stal'],"%Y-%m-%d")
            svarshik2 = Svarshik.objects.get_or_create(fio=data['svarshik_stal'],
                     type='Сталь', diametr=data['diametr_stal_styki'],
                     kolvo=data['kolvo_stal_styki'],date_svarki=date_stal)[0]
        else:
            svarshik2 = Svarshik.objects.get_or_create(fio=data['svarshik_stal'],
                     type='Сталь', diametr=data['diametr_stal_styki'],
                     kolvo=data['kolvo_stal_styki'])[0]
        svarshik2.save()
        object.svarshik2 = svarshik2

        if data['select_tehnadzor'] != "Выберите технадзор":
            person = Person.objects.get(fio=data['select_tehnadzor'])
            tehnadzor = TehNadzor.objects.get(person=person)
            object.tehnadzor = tehnadzor
        else:
            person = Person.objects.get_or_create(fio=data['select_tehnadzor_custom'])[0]
            person.save()
            tehnadzor = TehNadzor.objects.get_or_create(person=person)[0]
            tehnadzor.save()
            object.tehnadzor = tehnadzor

        if data['Date_1_zamer']:
            object.Data_zamera1 = datetime.datetime.strptime(data['Date_1_zamer'],"%Y-%m-%d")

        if data['Date_2_zamer']:
            object.Data_zamera2 = datetime.datetime.strptime(data['Date_2_zamer'],"%Y-%m-%d")

        if data['Date_sostavl_project']:
            object.Data_sost_project = datetime.datetime.strptime(data['Date_sostavl_project'],"%Y-%m-%d")

        if data['proektnaya_org']:
            object.proektnaya_org = data['proektnaya_org']

        if data['date_razbiv']:
            object.Data_razbiv = datetime.datetime.strptime(data['date_razbiv'],"%Y-%m-%d")

        if data['date_produv']:
            object.Data_produv = datetime.datetime.strptime(data['date_produv'],"%Y-%m-%d")

        if data['date_ukl']:
            object.Data_ukl = datetime.datetime.strptime(data['date_ukl'],"%Y-%m-%d")


        object.zashitnyi_futlyar = Diametr_x_dlina.objects.get_or_create(diametr=data['zashit_futlyar_diametr'],\
                                                                        x=data['zashit_futlyar_x'],\
                                                                        dlina=data['zashit_futlyar_diametr'])[0]

        object.futlyar_na_vyhode = Diametr_x_dlina_kolvo.objects.get_or_create(diametr=data['futlyar_na_vyhode_diametr'],\
                                                                        x=data['futlyar_na_vyhode_x'],\
                                                                        dlina = data['futlyar_na_vyhode_dlina'],\
                                                                        kolvo = data['futlyar_na_vyhode_kolvo'])[0]

        object.opora = Diametr_x_dlina_kolvo.objects.get_or_create(diametr=data['opora_pod_gazoprovod_diametr'],\
                                                                        x=data['opora_pod_gazoprovod_x'],\
                                                                        dlina = data['opora_pod_gazoprovod_dlina'],\
                                                                        kolvo = data['opora_pod_gazoprovod_kolvo'])[0]

        """ Доп данные шаровый кран"""

        elems = Dop_dann_sharovyi_kran.objects.all()


        if 'dop_dann_sharovyi_kran_mufta' in keys:
            dop_dann_sharovyi_kran = Dop_dann_sharovyi_kran()

            _add_dann_in_shar_kran(data=data,object=dop_dann_sharovyi_kran)
            dop_dann_sharovyi_kran.save()

            if elems:
                flag = False
                for elem in elems:
                    if (elem.pk != dop_dann_sharovyi_kran.pk and elem.kran == dop_dann_sharovyi_kran.kran
                        and elem.mufta == dop_dann_sharovyi_kran.mufta):
                        object.dop_dann_sharovyi_kran.add(elem)
                        dop_dann_sharovyi_kran.delete()
                        flag = True
                        break

                if not flag:
                    object.dop_dann_sharovyi_kran.add(dop_dann_sharovyi_kran)
            else:
                object.dop_dann_sharovyi_kran.add(dop_dann_sharovyi_kran)


        elems = Dop_dann_sharovyi_kran.objects.all()

        if indexes_dop_dann_shar_kran:
            for index in indexes_dop_dann_shar_kran:
                dop_dann_sharovyi_kran = Dop_dann_sharovyi_kran()
                _add_dann_in_shar_kran(data=data,object=dop_dann_sharovyi_kran,index='-{}'.format(index))
                dop_dann_sharovyi_kran.save()

                if elems:

                    flag = False
                    for elem in elems:
                        if (elem.pk != dop_dann_sharovyi_kran.pk and elem.kran == dop_dann_sharovyi_kran.kran
                            and elem.mufta == dop_dann_sharovyi_kran.mufta):
                            object.dop_dann_sharovyi_kran.add(elem)
                            dop_dann_sharovyi_kran.delete()
                            flag = True
                            break

                    if not flag:
                        object.dop_dann_sharovyi_kran.add(dop_dann_sharovyi_kran)

                else:
                    object.dop_dann_sharovyi_kran.add(dop_dann_sharovyi_kran)

        """ Доп данные газопровод в защитном футляре """

        elems = Dop_dann_gazopr_v_zashit.objects.all()

        if 'dop_dann_gazopr_v_zashit_diametr' in keys:
            dop_dann_gazopr_v_zashit = Dop_dann_gazopr_v_zashit()
            _add_dann_in_gazopr_v_zashit(data=data,object=dop_dann_gazopr_v_zashit)
            dop_dann_gazopr_v_zashit.save()

            if elems:

                flag = False
                for elem in elems:
                    if (elem.pk != dop_dann_gazopr_v_zashit.pk and elem.truba2 == dop_dann_gazopr_v_zashit.truba2
                        and elem.truba1 == dop_dann_gazopr_v_zashit.truba1):
                        object.dop_dann_gazopr_v_zashit.add(elem)
                        dop_dann_gazopr_v_zashit.delete()
                        flag = True
                        break

                if not flag:
                    object.dop_dann_gazopr_v_zashit.add(dop_dann_gazopr_v_zashit)
            else:
                object.dop_dann_gazopr_v_zashit.add(dop_dann_gazopr_v_zashit)


        elems = Dop_dann_gazopr_v_zashit.objects.all()

        if indexes_dop_dann_gazopr_v_zashit:
            for index in indexes_dop_dann_gazopr_v_zashit:
                dop_dann_gazopr_v_zashit = Dop_dann_gazopr_v_zashit()
                _add_dann_in_gazopr_v_zashit(data=data,object=dop_dann_gazopr_v_zashit,index='-{}'.format(index))
                dop_dann_gazopr_v_zashit.save()

                if elems:

                    flag = False

                    for elem in elems:
                        if (elem.pk != dop_dann_gazopr_v_zashit.pk and elem.truba2 == dop_dann_gazopr_v_zashit.truba2
                            and elem.truba1 == dop_dann_gazopr_v_zashit.truba1):
                            object.dop_dann_gazopr_v_zashit.add(elem)
                            dop_dann_gazopr_v_zashit.delete()
                            flag = True
                            break

                    if not flag:
                        object.dop_dann_gazopr_v_zashit.add(dop_dann_gazopr_v_zashit)

                else:
                    object.dop_dann_gazopr_v_zashit.add(dop_dann_gazopr_v_zashit)


        """ Доп данные об установке футляра """

        if 'dop_dann_ob_ustanovke_futlyarov_truba_diametr' in keys:
            dop_dann_ob_ustanovke_futl = Diametr_x.objects.get_or_create(\
                    diametr=data['dop_dann_ob_ustanovke_futlyarov_truba_diametr'],\
                    x=data['dop_dann_ob_ustanovke_futlyarov_truba_diametr'])[0]
            dop_dann_ob_ustanovke_futl.save()
            object.dop_dann_futlyar_na_vyhode.add(dop_dann_ob_ustanovke_futl)

        if indexes_dop_dann_ob_ustanovke_futl:
            for index in indexes_dop_dann_ob_ustanovke_futl:
                dop_dann_ob_ustanovke_futl = Diametr_x.objects.get_or_create(\
                    diametr=data['dop_dann_ob_ustanovke_futlyarov_truba_diametr'+'-{}'.format(index)],\
                    x=data['dop_dann_ob_ustanovke_futlyarov_truba_diametr'+'-{}'.format(index)])[0]
                dop_dann_ob_ustanovke_futl.save()
                object.dop_dann_futlyar_na_vyhode.add(dop_dann_ob_ustanovke_futl)

        """ Доп данные об установке опор """

        if 'dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr' in keys:
            dop_dann_ob_ustanovke_opor = Diametr_x.objects.get_or_create(\
                    diametr=data['dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr'],\
                    x=data['dop_dann_ob_ustanovke_opor_pod_gaz_truba_kolvo'])[0]
            dop_dann_ob_ustanovke_opor.save()
            object.dop_dann_ob_ustanovke_opor.add(dop_dann_ob_ustanovke_opor)

        if indexes_dop_dann_opor_pod:
            for index in indexes_dop_dann_opor_pod:
                dop_dann_ob_ustanovke_opor = Diametr_x.objects.get_or_create(\
                        diametr=data['dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr'+'-{}'.format(index)],\
                        x=data['dop_dann_ob_ustanovke_opor_pod_gaz_truba_kolvo'+'-{}'.format(index)])[0]
                dop_dann_ob_ustanovke_opor.save()
                object.dop_dann_ob_ustanovke_opor.add(dop_dann_ob_ustanovke_opor)

        """
        Доп. данные для Ростехнадзора
        """

        if 'pk1' in data.keys():
            rostehnadzor = ''
            rostehnadzor_truba = ''

            if data['rosteh_diametr'] or data['rosteh_x']:
                rostehnadzor_truba = Diametr_x.objects.get_or_create(
                    diametr=data['rosteh_diametr'], x=data['rosteh_x'])[0]
                rostehnadzor_truba.save()

            if data['pk1'] or data['pk1_diam'] or data['pk2'] or data['pk2_diam']:
                if rostehnadzor_truba:
                    rostehnadzor = Rostehnadzor.objects.get_or_create(
                        pk1=data['pk1'], pk1_diam=data['pk1_diam'],
                        pk2=data['pk2'], pk2_diam=data['pk2_diam'],
                        truba=rostehnadzor_truba)[0]
                else:
                    rostehnadzor = Rostehnadzor.objects.get_or_create(
                        pk1=data['pk1'], pk1_diam=data['pk1_diam'],
                        pk2=data['pk2'], pk2_diam=data['pk2_diam'])[0]

            else:
                if rostehnadzor_truba:
                    rostehnadzor = Rostehnadzor.objects.get_or_create(truba=rostehnadzor_truba)[0]

            if rostehnadzor:
                rostehnadzor.save()
                object.rostehnadzor = rostehnadzor

        """
        Доп. данные для Продавливания
        """

        if 'pk0_1_diam' in data.keys():
            prodavl = ''
            prodavl_truba = ''

            if data['prodav_diametr'] or data['prodav_x'] or data['prodav_dlina']:
                prodavl_truba = Diametr_x_dlina.objects.get_or_create(
                    diametr=data['prodav_diametr'], x=data['prodav_x'],
                    dlina=data['prodav_dlina'])[0]
                prodavl_truba.save()

            if data['pk0_1_diam'] or data['pk0_2_diam']:
                if prodavl_truba:
                    prodavl = Prodavlivanie.objects.get_or_create(
                        pk0_1_diam=data['pk0_1_diam'], pk0_2_diam=data['pk0_2_diam'],
                        truba=prodavl_truba)[0]
                else:
                    prodavl = Prodavlivanie.objects.get_or_create(
                        pk0_1_diam=data['pk0_1_diam'], pk0_2_diam=data['pk0_2_diam'])[0]

            else:
                if prodavl_truba:
                    prodavl = Prodavlivanie.objects.get_or_create(truba=prodavl_truba)[0]

            if prodavl:
                prodavl.save()
                object.prodavl = prodavl

        object.save()

        context['success'] = 'Объект успешно добавлен!'
        return render(request, 'object_add.html', {'context':context})


def form_document(request):
    context = {}
    objects = Object.objects.all().order_by('name_object')
    context['objects'] = objects

    if request.method == 'GET':

        return render(request, 'form_document.html', {'context':context})

    else:
        data = request.POST

        if data['select_object'] == 'Выберите объект' and data['select_field'] == 'Выберите тип документа':
            return render(request, 'form_document.html', {'context':context, 'error':'Выберите объект и тип документа!'})
        if data['select_object'] == 'Выберите объект':
            return render(request, 'form_document.html', {'context':context, 'error':'Выберите объект!'})

        object = Object.objects.get(name_object=data['select_object'])
        context['object'] = object

        if data['select_field'] == 'Выберите тип документа':
            return render(request, 'form_document.html', {'context':context, 'error':'Выберите тип документа!', 'selected':True})

        if data['select_field'] == 'ИТД':
            return form_ITD(data['select_object'])

        elif data['select_field'] == 'КС2 КС3 смета':
            if ('document' in request.FILES.keys()) and ('document_smeta' in request.FILES.keys()):
                uploaded_file = request.FILES['document']
                uploaded_file_smeta = request.FILES['document_smeta']
                return form_ks2(request, data['select_object'], uploaded_file, uploaded_file_smeta)
            else:
                return render(request, 'form_document.html', {'context':context, 'error':'Загрузите два файла для формирования документа!','selected':True})

        elif data['select_field'] == 'КС11 КС14':
            return form_ks14(data['select_object'])
        else:
            return render(request, 'form_document.html', {'context':context, 'error':'Ошибка!','selected':True})


def object_smeta(request):
    context = {}
    context['certificates'] = Certificate.objects.all().order_by('name')
    context['objects'] = Object.objects.all().order_by('name_object')

    if request.method == 'GET':
        return render(request, 'object_smeta.html', {'context':context})
    
    if request.method == 'POST':
        data = request.POST
        if data['select_object'] == 'Выберите объект':
            return render(request, 'object_smeta.html', {'context':context, 'error':'Выберите объект!'})
        
        object = Object.objects.get(name_object=data['select_object'])
        if object.smeta:
            if object.smeta.date_ks2:
                object.smeta.date_ks2 = datetime.datetime.strftime(object.smeta.date_ks2,"%Y-%m-%d")
            if object.smeta.date_nach_rabot:
                object.smeta.date_nach_rabot = datetime.datetime.strftime(object.smeta.date_nach_rabot,"%Y-%m")
            if object.smeta.date_nach_zakr:
                object.smeta.date_nach_zakr = datetime.datetime.strftime(object.smeta.date_nach_zakr,"%Y-%m-%d")
            if object.smeta.date_kon_zakr:
                object.smeta.date_kon_zakr = datetime.datetime.strftime(object.smeta.date_kon_zakr,"%Y-%m-%d")
            if object.smeta.date_dogovor:
                object.smeta.date_dogovor = datetime.datetime.strftime(object.smeta.date_dogovor,"%Y-%m-%d")
        
        context['object'] = object
        
        return render(request, 'object_smeta.html', {'context':context, 'selected':True})


def object_add_data_smeta(request):
    context = {}
    context['objects'] = Object.objects.all().order_by('name_object')

    data = request.POST
    object = Object.objects.get(name_object=data['select_object'])

    smeta = Smeta()

    if data['summa_proektnoi_smeti']:
        smeta.summa_proektnoi_smeti = data['summa_proektnoi_smeti']

    if data['summa_utv_smeti']:
        smeta.summa_utv_smeti = data['summa_utv_smeti']

    if data['summa_ks2_bez_nds']:
        smeta.summa_ks2_bez_nds = data['summa_ks2_bez_nds']

    if data['date_ks2']:
        smeta.date_ks2 = data['date_ks2']

    if data['date_nach_rabot']:
        smeta.date_nach_rabot = data['date_nach_rabot']

    if data['date_nach_zakr']:
        smeta.date_nach_zakr = data['date_nach_zakr']

    if data['date_kon_zakr']:
        smeta.date_kon_zakr = data['date_kon_zakr']

    if data['nomer_dogovor']:
        smeta.nomer_dogovor = data['nomer_dogovor']

    if data['date_dogovor']:
        smeta.date_dogovor = data['date_dogovor']

    if data['date_nach_rabot']:
        date_nach_rabot = data['date_nach_rabot']
        full_date_nach_rabot = f"{date_nach_rabot}-01"
        smeta.date_nach_rabot =  full_date_nach_rabot

    smeta.save()
    object.smeta = smeta
    object.save()

    return render(request, 'object_smeta.html', {'context':context, 'success':'Данные для сметы успешно добавлены!'})


def object_select_view(request):
    if request.method == 'GET':
        context = {}
        objects = Object.objects.all().order_by('name_object')
        context['objects'] = objects
        context['fields'] = fields
        return render(request, 'object_view.html', {'context':context})

    else:
        context = {}

        data = request.POST

        objects = Object.objects.all().order_by('name_object')
        context['objects'] = objects
        context['fields'] = fields

        if 'delete' in data.keys() and data['select_object'] == 'Выберите объект':
            return render(request, 'object_view.html', {'context':context, 'error':'Выберите объект!'})

        if data['select_object'] == 'Выберите объект':
            return render(request, 'object_view.html', {'context':context, 'error':'Выберите объект!'})

        if 'delete' in data.keys():
            object = Object.objects.get(name_object=data['select_object'])
            object.delete()
            return render(request, 'object_view.html', {'context':context, 'success':'Объект успешно удалён!'})

        if data['select_object'] == 'Выберите объект':
            if 'select_field' in request.POST.keys():
                return render(request, 'object_view.html', {'context':context, 'error':'Выберите объект!'})
            else:
                return render(request, 'object_view.html', {'context':context, 'error':'Выберите объект и данные!'})

        if 'select_field' not in request.POST.keys():
            return render(request, 'object_view.html', {'context':context, 'error':'Выберите данные!'})

        object = Object.objects.get(name_object=data['select_object'])

        fields_db = data.getlist('select_field')

        fields_names = [fields[field_db] for field_db in fields_db]

        context['object'] = object

        context['field_db'] = fields_db
        context['field_name'] = fields_names
        context['certificates'] = Certificate.objects.all().order_by('name')
        context['field_db_names'] = list(zip(fields_db,fields_names))

        context['field_value'] = []

        for field_db in fields_db:
            try:
                if field_db in fields_dates:
                    if getattr(object,field_db):
                        context['field_value'].append(
                        datetime.datetime.strftime(getattr(object,field_db), "%Y-%m-%d"))
                    else:
                        context['field_value'].append('')
                else:
                    context['field_value'].append(getattr(object,field_db))
            except:
                context['field_value'].append('')

        context['field_db_names_values'] = list(zip(fields_db,fields_names,context['field_value']))

        context['fields'] = fields
        context['fields_dates'] = fields_dates

        context['dop_dann_sharovyi_krans'] = list(zip(range(len(object.dop_dann_sharovyi_kran.all())),
            object.dop_dann_sharovyi_kran.all()))
        context['dop_dann_gazopr_v_zashits'] = list(zip(range(len(object.dop_dann_gazopr_v_zashit.all())),
            object.dop_dann_gazopr_v_zashit.all()))
        context['dop_dann_futlyar_na_vyhodes'] = list(zip(range(len(object.dop_dann_futlyar_na_vyhode.all())),
            object.dop_dann_futlyar_na_vyhode.all()))
        context['dop_dann_ob_ustanovke_opors'] = list(zip(range(len(object.dop_dann_ob_ustanovke_opor.all())),
            object.dop_dann_ob_ustanovke_opor.all()))

        context['indexes_podzem_stal'] = {}
        index = 0
        itogo = 0
        if object.gazoprovod_podzem_stal:

            if object.gazoprovod_podzem_stal.truba.all():
                for item in object.gazoprovod_podzem_stal.truba.all():
                    context['indexes_podzem_stal'][index] = {}
                    context['indexes_podzem_stal'][index]['truba'] = item
                    try:
                        itogo += float(item.dlina)
                    except:
                        itogo += 0

                    index += 1
        context['indexes_podzem_stal']['object'] = object.gazoprovod_podzem_stal
        context['indexes_podzem_stal']['itogo'] = itogo

        context['indexes_podzem_poliet'] = {}
        index = 0
        itogo = 0
        if object.gazoprovod_podzem_poliet:
            if object.gazoprovod_podzem_poliet.truba.all():
                for item in object.gazoprovod_podzem_poliet.truba.all():
                    context['indexes_podzem_poliet'][index] = {}
                    context['indexes_podzem_poliet'][index]['truba'] = item
                    index += 1
                    try:
                        itogo += float(item.dlina)
                    except:
                        itogo += 0
        index = 0
        if object.gazoprovod_podzem_poliet:
            if object.gazoprovod_podzem_poliet.mufta.all():
                for item in object.gazoprovod_podzem_poliet.mufta.all():
                    context['indexes_podzem_poliet'][index]['mufta'] = item
                    index += 1
                    try:
                        itogo += float(item.dlina)
                    except:
                        itogo += 0

        context['indexes_podzem_poliet']['object'] = object.gazoprovod_podzem_poliet
        context['indexes_podzem_poliet']['itogo'] = itogo

        context['indexes_nadzem_stal'] = {}
        index = 0
        itogo = 0
        if object.gazoprovod_nadzem_stal:
            if object.gazoprovod_nadzem_stal.truba.all():
                for item in object.gazoprovod_nadzem_stal.truba.all():
                    context['indexes_nadzem_stal'][index] = {}
                    context['indexes_nadzem_stal'][index]['truba'] = item
                    index += 1
                    try:
                        itogo += float(item.dlina)
                    except:
                        itogo += 0
        context['indexes_nadzem_stal']['object'] = object.gazoprovod_nadzem_stal
        context['indexes_nadzem_stal']['itogo'] = itogo

        context['indexes_dop_dann_sharovyi_kran'] = {}
        index = 0
        if object.dop_dann_sharovyi_kran.all():
            for item in object.dop_dann_sharovyi_kran.all():
                context['indexes_dop_dann_sharovyi_kran'][index] = {}
                context['indexes_dop_dann_sharovyi_kran'][index]['kran'] = item.kran
                context['indexes_dop_dann_sharovyi_kran'][index]['mufta'] = item.mufta
                index += 1


        context['indexes_dop_dann_gazopr_v_zashit'] = {}
        index = 0
        if object.dop_dann_gazopr_v_zashit.all():
            for item in object.dop_dann_gazopr_v_zashit.all():
                context['indexes_dop_dann_gazopr_v_zashit'][index] = {}
                context['indexes_dop_dann_gazopr_v_zashit'][index]['truba1'] = item.truba1
                context['indexes_dop_dann_gazopr_v_zashit'][index]['truba2'] = item.truba2
                index += 1

        context['indexes_dop_dann_futlyar_na_vyhode'] = {}
        index = 0
        if object.dop_dann_futlyar_na_vyhode.all():
            for item in object.dop_dann_futlyar_na_vyhode.all():
                context['indexes_dop_dann_futlyar_na_vyhode'][index] = item
                index += 1

        context['indexes_dop_dann_ob_ustanovke_opor'] = {}
        index = 0
        if object.dop_dann_ob_ustanovke_opor.all():
            for item in object.dop_dann_ob_ustanovke_opor.all():
                context['indexes_dop_dann_ob_ustanovke_opor'][index] = item
                index += 1

        return render(request, 'object_view.html', {'context':context, 'selected':True})


def object_update(request):
    data = request.POST

    rostehnadzor_fields = ('pk1', 'pk1_diam', 'pk2', 'pk2_diam', 'rosteh_diametr', 'rosteh_x')
    prodavl_fields = ('pk0_1_diam', 'pk0_2_diam', 'prodav_diametr', 'prodav_x', 'prodav_dlina')

    kontragent_fields = ('name_kontragent', 'INN', 'KPP', 'Ur_address',
        'telephone', 'ks11_predstav_podryadchika_fio',
        'ks11_predstav_podryadchika_post', 'podpisant')

    gip_fields = ('custom_GIP', 'organization')

    ks11_fields = ('ks11_predsedatel_fio', 'ks11_predsedatel_post',
        'ks11_predstav_proekt_fio', 'ks11_predstav_proekt_post',
        'ks11_predstav_ekspl_fio', 'ks11_predstav_ekspl_post',
        'ks11_predstav_podryadchika_post','ks11_predstav_podryadchika_fio')

    ks2_fields = ('ks2_zakazchik_post', 'ks2_zakazchik_fio', 'ks2_podryadchik_post',
        'ks2_podryadchik_fio', 'object', 'stroika')

    svarshik1_fields = ('svarshik1_fio', 'svarshik1_diametr', 'svarshik1_kolvo',
        'svarshik1_date_svarki')

    svarshik2_fields = ('svarshik2_fio', 'svarshik2_diametr', 'svarshik2_kolvo',
        'svarshik2_date_svarki')

    zashit_futlyar_fields = ('zashit_futlyar_diametr', 'zashit_futlyar_x',
        'zashit_futlyar_dlina')

    futlyar_na_vyhode_fields = ('futlyar_na_vyhode_diametr', 'futlyar_na_vyhode_x',
        'futlyar_na_vyhode_dlina', 'futlyar_na_vyhode_kolvo')

    opora_pod_gazoprovod_fields = ('opora_pod_gazoprovod_diametr', 'opora_pod_gazoprovod_x',
        'opora_pod_gazoprovod_dlina', 'opora_pod_gazoprovod_kolvo')

    smeta_fields = ('summa_proektnoi_smeti', 'summa_utv_smeti', 'summa_ks2_bez_nds',
        'date_ks2', 'date_nach_rabot', 'date_nach_zakr', 'date_kon_zakr', 'nomer_dogovor',
        'date_dogovor')

    podzem_stal_gaz_fields = ('stalnoi_neraz_soed_PE', 'stalnoi_neraz_soed_ST',
        'neraz_soed_kolvo', 'stalnoi_kontrolnaya_trubka', 'stalnoi_otvod',
        'stal_opoznavat_znak')

    podzem_poliet_gaz_fields = ('poliet_otvod_diametr', 'poliet_otvod_dlina',
        'poliet_troinik1', 'poliet_troinik2', 'poliet_troinik3',
        'poliet_troinik_dlina', 'poliet_zaglushka_diametr',
        'poliet_zaglushka_dlina', 'poliet_lenta_signal_dlina',
        'poliet_kran_shar_diametr', 'poliet_kran_shar_dlina',
        'poliet_opoznavat_znak_kolvo', 'sedelka_poliet_troinik1',
        'sedelka_poliet_troinik2', 'sedelka_poliet_troinik3',
        'sedelka_poliet_kolvo')

    nadzem_gaz_fields = ('nadzem_stal_izolir_soed', 'nadzem_stal_izolir_kolvo',
        'nadzem_stal_kran_diametr', 'nadzem_stal_kran_kolvo', 'nadzem_stal_otvod_diametr',
        'nadzem_stal_otvod_kolvo', 'nadzem_stal_cokol_vvod_PE',
        'nadzem_stal_cokol_vvod_ST', 'nadzem_stal_cokol_vvod_kolvo',
        'nadzem_stal_kreplenie', 'nadzem_stalnoi_stoika_diametr', 'nadzem_stalnoi_stoika_dlina',
        'nadzem_stalnoi_stoika_kolvo')

    object = Object.objects.get(name_object=data['obj'])

    if 'gazoprovod' in data.keys():
        object.gazoprovod = True
    else:
        object.gazoprovod = False

    if 'prodavlivanie' in data.keys():
        object.prodavlivanie = True
    else:
        object.prodavlivanie = False

    if 'select_cert' in data.keys():
        object.certificates.clear()
        for cert in data.getlist('select_cert'):
            certificate = Certificate.objects.get(name=cert)
            object.certificates.add(certificate)

    existed_new_gip = []
    existed_person_podryadchik = []
    existed_person_preds = []
    existed_person_predstav_eks = []
    existed_person_predstav_proek = []
    existed_new_kontragent = []
    existed_new_podpisant = []
    existed_person_zakazchik = []
    existed_person_ks2_podryadchik = []
    existed_new_svarshik1 = []
    existed_new_svarshik2 = []
    existed_new_zashit_futlyar = []
    existed_new_opora = []
    existed_new_futlyar_na_vyh = []
    existed_new_smeta = []
    existed_new_tehnadzor = []
    existed_new_tehnadzor_person = []
    existed_new_podzem_stal_gazoprovod = []
    existed_new_neraz_soed_truba_podzem_stal_gazoprovod = []
    existed_new_podzem_poliet_gazoprovod = []
    existed_new_otvod = []
    existed_new_troinik = []
    existed_new_sedelka = []
    existed_new_zaglushka = []
    existed_new_kran = []
    existed_new_nadzem_stal_gazoprovod = []
    existed_new_nadzem_stal_gazoprovod_izolir = []
    existed_new_nadzem_stal_gazoprovod_kran = []
    existed_new_nadzem_stal_gazoprovod_otvod = []
    existed_new_nadzem_stal_gazoprovod_cokol = []
    existed_new_stoika = []
    existed_new_rostehnadzor = []
    existed_new_rostehnadzor_truba = []
    existed_new_prodavl = []
    existed_new_prodavl_truba = []

    new_gip = ''
    person_podryadchik = ''
    person_preds = ''
    person_predstav_eks = ''
    person_predstav_proek = ''
    new_kontragent = ''
    new_podpisant = ''
    person_zakazchik = ''
    person_ks2_podryadchik = ''
    new_svarshik1 = ''
    new_svarshik2 = ''
    new_zashit_futlyar = ''
    new_opora = ''
    new_futlyar_na_vyh = ''
    new_smeta = ''
    new_tehnadzor = ''
    new_tehnadzor_person = ''
    new_podzem_stal_gazoprovod = ''
    new_neraz_soed_truba_podzem_stal_gazoprovod = ''
    new_podzem_poliet_gazoprovod = ''
    new_otvod = ''
    new_troinik = ''
    new_sedelka = ''
    new_zaglushka = ''
    new_kran = ''
    new_nadzem_stal_gazoprovod = ''
    new_nadzem_stal_gazoprovod_izolir = ''
    new_nadzem_stal_gazoprovod_kran = ''
    new_nadzem_stal_gazoprovod_otvod = ''
    new_nadzem_stal_gazoprovod_cokol = ''
    new_stoika = ''
    new_district = ''
    new_rostehnadzor = ''
    new_prodavl = ''
    new_rostehnadzor_truba = ''
    new_prodavl_truba = ''
    flag_teh = False

    for field in data:
        if field in gip_fields:
            new_gip = GIP()
        elif field in ks11_fields:
            person_podryadchik = Person()
            person_preds = Person()
            person_predstav_eks = Person()
            person_predstav_proek = Person()
        elif field in kontragent_fields:
            new_kontragent = Kontragent()
            new_podpisant = Person()
        elif field in ks2_fields:
            person_zakazchik = Person()
            person_ks2_podryadchik = Person()
        elif field in svarshik1_fields:
            new_svarshik1 = Svarshik()
        elif field in svarshik2_fields:
            new_svarshik2 = Svarshik()
        elif field in zashit_futlyar_fields:
            new_zashit_futlyar = Diametr_x_dlina()
        elif field in opora_pod_gazoprovod_fields:
            new_opora = Diametr_x_dlina_kolvo()
        elif field in futlyar_na_vyhode_fields:
            new_futlyar_na_vyh = Diametr_x_dlina_kolvo()
        elif field in smeta_fields:
            new_smeta = Smeta()
        elif field == 'tehnadzor_fio':
            new_tehnadzor = TehNadzor()
            new_tehnadzor_person = Person()
        elif field in podzem_stal_gaz_fields:
            new_podzem_stal_gazoprovod = Podzem_stal_gazoprovod()
            new_podzem_stal_gazoprovod.save()
            new_neraz_soed_truba_podzem_stal_gazoprovod = Neraz_soed_stal_gazoprovod()
        elif field in podzem_poliet_gaz_fields:
            new_podzem_poliet_gazoprovod = Podzem_polietilen_gazoprovod()
            new_podzem_poliet_gazoprovod.save()
            new_otvod = Diametr_kolvo()
            new_troinik = Diametrs_3_kolvo()
            new_sedelka = Diametrs_3_kolvo()
            new_zaglushka = Diametr_kolvo()
            new_kran = Diametr_kolvo()
        elif field in nadzem_gaz_fields:
            new_nadzem_stal_gazoprovod = Nadzem_stal_gazoprovod()
            new_nadzem_stal_gazoprovod.save()
            new_nadzem_stal_gazoprovod_izolir = Diametr_kolvo()
            new_nadzem_stal_gazoprovod_kran = Diametr_kolvo()
            new_nadzem_stal_gazoprovod_otvod = Diametr_kolvo()
            new_nadzem_stal_gazoprovod_cokol = Cokol_soed_stal_gazoprovod()
            new_stoika = Diametr_dlina_kolvo()
        elif field in rostehnadzor_fields:
            new_rostehnadzor = Rostehnadzor()
            new_rostehnadzor_truba = Diametr_x()

        elif field in prodavl_fields:
            new_prodavl = Prodavlivanie()
            new_prodavl_truba = Diametr_x_dlina()

    len_stal_truba = 0
    if 'stalnoi_truba_diametr' in data.keys():
        len_stal_truba = len(data.getlist('stalnoi_truba_diametr'))

    if len_stal_truba:
        for index in range(len_stal_truba):
            new_truba_podzem_stal_gazoprovod = Diametr_x_dlina_prim.objects.get_or_create(
                                                diametr = data.getlist('stalnoi_truba_diametr')[index],\
                                                x = data.getlist('stalnoi_truba_kolvo')[index], \
                                                dlina = data.getlist('stalnoi_truba_dlina')[index],\
                                                prim = data.getlist('stalnoi_truba_prim')[index])[0]

            new_truba_podzem_stal_gazoprovod.save()
            new_podzem_stal_gazoprovod.truba.add(new_truba_podzem_stal_gazoprovod)
            new_podzem_stal_gazoprovod.save()
            object.gazoprovod_podzem_stal = new_podzem_stal_gazoprovod

    len_poliet_truba = 0
    if 'poliet_truba_diametr' in data.keys():
        len_poliet_truba = len(data.getlist('poliet_truba_diametr'))

    if len_poliet_truba:
        for index in range(len_poliet_truba):
            new_truba_podzem_poliet_gazoprovod = Diametr_x_dlina.objects.get_or_create(
                                                diametr = data.getlist('poliet_truba_diametr')[index],\
                                                x = data.getlist('poliet_truba_kolvo')[index], \
                                                dlina = data.getlist('poliet_truba_dlina')[index])[0]

            new_truba_podzem_poliet_gazoprovod.save()

            new_mufta_podzem_poliet_gazoprovod = Diametr_x_dlina.objects.get_or_create(
                                                diametr = data.getlist('poliet_mufta_diametr')[index],\
                                                x = data.getlist('poliet_mufta_kolvo')[index], \
                                                dlina = data.getlist('poliet_mufta_dlina')[index])[0]
            new_mufta_podzem_poliet_gazoprovod.save()
            new_podzem_poliet_gazoprovod.truba.add(new_truba_podzem_poliet_gazoprovod)
            new_podzem_poliet_gazoprovod.mufta.add(new_mufta_podzem_poliet_gazoprovod)
            new_podzem_poliet_gazoprovod.save()
            object.gazoprovod_podzem_poliet = new_podzem_poliet_gazoprovod

    len_nadzem_truba = 0
    if 'nadzem_stal_truba_diametr' in data.keys():
        len_nadzem_truba = len(data.getlist('nadzem_stal_truba_diametr'))

    if len_nadzem_truba:
        for index in range(len_nadzem_truba):
            new_truba_nadzem_gazoprovod = Diametr_x_dlina.objects.get_or_create(
                                                diametr = data.getlist('nadzem_stal_truba_diametr')[index],\
                                                x = data.getlist('nadzem_stal_truba_x')[index], \
                                                dlina = data.getlist('nadzem_stal_truba_dlina')[index])[0]

            new_truba_nadzem_gazoprovod.save()

            new_nadzem_stal_gazoprovod.truba.add(new_truba_nadzem_gazoprovod)

    len_dop_dann_shar = 0
    if 'dop_dann_sharovyi_kran' in data.keys():
        len_dop_dann_shar = len(data.getlist('dop_dann_sharovyi_kran'))
        object.dop_dann_sharovyi_kran.clear()

    if len_dop_dann_shar:
        for index in range(len_dop_dann_shar):

            diam1 = Diametr.objects.get_or_create(diametr=data.getlist('dop_dann_sharovyi_kran')[index])[0]
            diam1.save()

            diam2 = Diametr.objects.get_or_create(diametr=data.getlist('dop_dann_sharovyi_kran_mufta')[index])[0]
            diam2.save()

            new_dop_dann_sharovyi_kran = Dop_dann_sharovyi_kran.objects.get_or_create(
                kran=diam1, mufta=diam2)[0]
            new_dop_dann_sharovyi_kran.save()

            object.dop_dann_sharovyi_kran.add(new_dop_dann_sharovyi_kran)

    len_dop_dann_gazopr_v_zashit = 0
    if 'dop_dann_gazopr_v_zashit_diametr' in data.keys():
        len_dop_dann_gazopr_v_zashit = len(data.getlist('dop_dann_gazopr_v_zashit_diametr'))
        object.dop_dann_gazopr_v_zashit.clear()

    if len_dop_dann_gazopr_v_zashit:
        for index in range(len_dop_dann_gazopr_v_zashit):

            truba1 = Diametr_x.objects.get_or_create(diametr=data.getlist('dop_dann_gazopr_v_zashit_diametr')[index],
                x=data.getlist('dop_dann_gazopr_v_zashit_kolvo')[index])[0]
            truba1.save()

            truba2= Diametr_x.objects.get_or_create(diametr=data.getlist('dop_dann_gazopr_v_zashit_truba_diametr')[index],
                x=data.getlist('dop_dann_gazopr_v_zashit_truba_kolvo')[index])[0]
            truba2.save()

            new_dop_dann_gazopr_v_zashit = Dop_dann_gazopr_v_zashit.objects.get_or_create(
                truba1=truba1, truba2=truba2)[0]
            new_dop_dann_gazopr_v_zashit.save()

            object.dop_dann_gazopr_v_zashit.add(new_dop_dann_gazopr_v_zashit)

    len_dop_dann_futlyar_na_vyhode = 0
    if 'dop_dann_ob_ustanovke_futlyarov_truba_diametr' in data.keys():
        len_dop_dann_futlyar_na_vyhode = len(data.getlist('dop_dann_ob_ustanovke_futlyarov_truba_diametr'))
        object.dop_dann_futlyar_na_vyhode.clear()

    if len_dop_dann_futlyar_na_vyhode:
        for index in range(len_dop_dann_futlyar_na_vyhode):

            truba1 = Diametr_x.objects.get_or_create(diametr=data.getlist('dop_dann_ob_ustanovke_futlyarov_truba_diametr')[index],
                x=data.getlist('dop_dann_ob_ustanovke_futlyarov_truba_kolvo')[index])[0]
            truba1.save()

            object.dop_dann_futlyar_na_vyhode.add(truba1)

    len_dop_dann_ob_ustanovke_opor = 0
    if 'dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr' in data.keys():
        len_dop_dann_ob_ustanovke_opor = len(data.getlist('dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr'))
        object.dop_dann_ob_ustanovke_opor.clear()

    if len_dop_dann_ob_ustanovke_opor:
        for index in range(len_dop_dann_ob_ustanovke_opor):

            truba1 = Diametr_x.objects.get_or_create(diametr=data.getlist('dop_dann_ob_ustanovke_opor_pod_gaz_truba_diametr')[index],
                x=data.getlist('dop_dann_ob_ustanovke_opor_pod_gaz_truba_kolvo')[index])[0]
            truba1.save()

            object.dop_dann_ob_ustanovke_opor.add(truba1)

    for field in data:

        if data[field]:

            if field in rostehnadzor_fields:

                if field == 'pk1':

                    if existed_new_rostehnadzor:
                        for teh in existed_new_rostehnadzor:
                            if teh.pk1 != data[field]:
                                existed_new_rostehnadzor.remove(teh)

                    if data[field]:

                        if Rostehnadzor.objects.filter(pk1=data[field]):
                            for elem in Rostehnadzor.objects.filter(pk1=data[field]):
                                existed_new_rostehnadzor.append(elem)

                    new_rostehnadzor.pk1 = data[field]

                elif field == 'pk1_diam':

                    if existed_new_rostehnadzor:
                        for teh in existed_new_rostehnadzor:
                            if teh.pk1_diam != data[field]:
                                existed_new_rostehnadzor.remove(teh)

                    if data[field]:

                        if Rostehnadzor.objects.filter(pk1_diam=data[field]):
                            for elem in Rostehnadzor.objects.filter(pk1_diam=data[field]):
                                existed_new_rostehnadzor.append(elem)

                    new_rostehnadzor.pk1_diam = data[field]

                elif field == 'pk2':

                    if existed_new_rostehnadzor:
                        for teh in existed_new_rostehnadzor:
                            if teh.pk2 != data[field]:
                                existed_new_rostehnadzor.remove(teh)

                    if data[field]:

                        if Rostehnadzor.objects.filter(pk2=data[field]):
                            for elem in Rostehnadzor.objects.filter(pk2=data[field]):
                                existed_new_rostehnadzor.append(elem)

                    new_rostehnadzor.pk2 = data[field]

                elif field == 'pk2_diam':

                    if existed_new_rostehnadzor:
                        for teh in existed_new_rostehnadzor:
                            if teh.pk2_diam != data[field]:
                                existed_new_rostehnadzor.remove(teh)

                    if data[field]:

                        if Rostehnadzor.objects.filter(pk2_diam=data[field]):
                            for elem in Rostehnadzor.objects.filter(pk2_diam=data[field]):
                                existed_new_rostehnadzor.append(elem)

                    new_rostehnadzor.pk2_diam = data[field]

                elif field == 'rosteh_diametr':

                    if existed_new_rostehnadzor_truba:
                        for teh in existed_new_rostehnadzor_truba:
                            if teh.diametr != data[field]:
                                existed_new_rostehnadzor_truba.remove(teh)

                    if data[field]:

                        if Diametr_x.objects.filter(diametr=data[field]):
                            for elem in Diametr_x.objects.filter(diametr=data[field]):
                                existed_new_rostehnadzor_truba.append(elem)

                    new_rostehnadzor_truba.diametr = data[field]

                else:

                    if existed_new_rostehnadzor_truba:
                        for teh in existed_new_rostehnadzor_truba:
                            if teh.x != data[field]:
                                existed_new_rostehnadzor_truba.remove(teh)

                    if data[field]:

                        if Diametr_x.objects.filter(x=data[field]):
                            for elem in Diametr_x.objects.filter(x=data[field]):
                                existed_new_rostehnadzor_truba.append(elem)

                    new_rostehnadzor_truba.x = data[field]

            elif field in prodavl_fields:

                if field == 'prodav_diametr':

                    if existed_new_prodavl_truba:
                        for teh in existed_new_prodavl_truba:
                            if teh.diametr != data[field]:
                                existed_new_prodavl_truba.remove(teh)

                    if data[field]:

                        if Diametr_x_dlina.objects.filter(diametr=data[field]):
                            for elem in Diametr_x_dlina.objects.filter(diametr=data[field]):
                                existed_new_prodavl_truba.append(elem)

                    new_prodavl_truba.diametr = data[field]

                if field == 'prodav_x':

                    if existed_new_prodavl_truba:
                        for teh in existed_new_prodavl_truba:
                            if teh.x != data[field]:
                                existed_new_prodavl_truba.remove(teh)

                    if data[field]:

                        if Diametr_x_dlina.objects.filter(x=data[field]):
                            for elem in Diametr_x_dlina.objects.filter(x=data[field]):
                                existed_new_prodavl_truba.append(elem)

                    new_prodavl_truba.x = data[field]

                if field == 'prodav_dlina':

                    if existed_new_prodavl_truba:
                        for teh in existed_new_prodavl_truba:
                            if teh.dlina != data[field]:
                                existed_new_prodavl_truba.remove(teh)

                    if data[field]:

                        if Diametr_x_dlina.objects.filter(dlina=data[field]):
                            for elem in Diametr_x_dlina.objects.filter(dlina=data[field]):
                                existed_new_prodavl_truba.append(elem)

                    new_prodavl_truba.dlina = data[field]

                if field == 'pk0_1_diam':

                    if existed_new_prodavl:
                        for teh in existed_new_prodavl:
                            if teh.pk0_1_diam != data[field]:
                                existed_new_prodavl.remove(teh)

                    if data[field]:

                        if Prodavlivanie.objects.filter(pk0_1_diam=data[field]):
                            for elem in Prodavlivanie.objects.filter(pk0_1_diam=data[field]):
                                existed_new_prodavl.append(elem)

                    new_prodavl.pk0_1_diam = data[field]

                if field == 'pk0_2_diam':

                    if existed_new_prodavl:
                        for teh in existed_new_prodavl:
                            if teh.pk0_1_diam != data[field]:
                                existed_new_prodavl.remove(teh)

                    if data[field]:

                        if Prodavlivanie.objects.filter(pk0_2_diam=data[field]):
                            for elem in Prodavlivanie.objects.filter(pk0_2_diam=data[field]):
                                existed_new_prodavl.append(elem)

                    new_prodavl.pk0_2_diam = data[field]

            elif field in gip_fields:

                if field == 'custom_GIP':

                    if existed_new_gip:
                        for gip in existed_new_gip:
                            if gip.fio != data[field]:
                                existed_new_gip.remove(gip)

                    if GIP.objects.filter(fio=data[field]):
                        for elem in GIP.objects.filter(fio=data[field]):
                            existed_new_gip.append(elem)

                    new_gip.fio = data[field]

                else:
                    if existed_new_gip:
                        for gip in existed_new_gip:
                            if gip.organization != data[field]:
                                existed_new_gip.remove(gip)

                    if GIP.objects.filter(organization=data[field]):
                        for elem in GIP.objects.filter(organization=data[field]):
                            existed_new_gip.append(elem)

                    new_gip.organization = data[field]


            elif field in nadzem_gaz_fields:

                if 'nadzem_stal_izolir' in field:

                    if existed_new_nadzem_stal_gazoprovod_izolir:
                        for gzp in existed_new_nadzem_stal_gazoprovod_izolir:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_nadzem_stal_gazoprovod_izolir.remove(gzp)

                    if field.split('_')[-1] == 'kolvo':
                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_nadzem_stal_gazoprovod_izolir.append(elem)

                        setattr(new_nadzem_stal_gazoprovod_izolir, 'kolvo', data[field])
                    else:
                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_nadzem_stal_gazoprovod_izolir.append(elem)

                        setattr(new_nadzem_stal_gazoprovod_izolir, 'diametr', data[field])

                if 'nadzem_stal_kran' in field:

                    if existed_new_nadzem_stal_gazoprovod_kran:
                        for gzp in existed_new_nadzem_stal_gazoprovod_kran:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_nadzem_stal_gazoprovod_kran.remove(gzp)

                    if 'diametr' in field:
                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_nadzem_stal_gazoprovod_kran.append(elem)

                        setattr(new_nadzem_stal_gazoprovod_kran, 'diametr', data[field])
                    else:
                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_nadzem_stal_gazoprovod_kran.append(elem)

                        setattr(new_nadzem_stal_gazoprovod_kran, 'kolvo', data[field])

                if 'nadzem_stal_otvod' in field:
                    if existed_new_nadzem_stal_gazoprovod_otvod:
                        for gzp in existed_new_nadzem_stal_gazoprovod_otvod:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_nadzem_stal_gazoprovod_otvod.remove(gzp)

                    if 'diametr' in field:
                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_nadzem_stal_gazoprovod_otvod.append(elem)
                        setattr(new_nadzem_stal_gazoprovod_otvod, 'diametr', data[field])
                    else:
                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_nadzem_stal_gazoprovod_otvod.append(elem)

                        setattr(new_nadzem_stal_gazoprovod_otvod, 'kolvo', data[field])

                if 'nadzem_stal_cokol_vvod' in field:
                    if existed_new_nadzem_stal_gazoprovod_cokol:
                        for gzp in existed_new_nadzem_stal_gazoprovod_cokol:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_nadzem_stal_gazoprovod_cokol.remove(gzp)

                    if 'PE' in field:
                        if Cokol_soed_stal_gazoprovod.objects.filter(PE=data[field]):
                            for elem in Cokol_soed_stal_gazoprovod.objects.filter(PE=data[field]):
                                existed_new_nadzem_stal_gazoprovod_cokol.append(elem)
                    elif 'ST' in field:
                        if Cokol_soed_stal_gazoprovod.objects.filter(ST=data[field]):
                            for elem in Cokol_soed_stal_gazoprovod.objects.filter(ST=data[field]):
                                existed_new_nadzem_stal_gazoprovod_cokol.append(elem)
                    else:
                        if Cokol_soed_stal_gazoprovod.objects.filter(kolvo=data[field]):
                            for elem in Cokol_soed_stal_gazoprovod.objects.filter(kolvo=data[field]):
                                existed_new_nadzem_stal_gazoprovod_cokol.append(elem)

                    setattr(new_nadzem_stal_gazoprovod_cokol, field.split('_')[-1], data[field])

                if 'nadzem_stalnoi_stoika' in field:
                    if existed_new_stoika:
                        for gzp in existed_new_stoika:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_stoika.remove(gzp)

                    if 'dlina' in field:
                        if Diametr_dlina_kolvo.objects.filter(dlina=data[field]):
                            for elem in Diametr_dlina_kolvo.objects.filter(dlina=data[field]):
                                existed_new_stoika.append(elem)
                    if 'diametr' in field:
                        if Diametr_dlina_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_dlina_kolvo.objects.filter(diametr=data[field]):
                                existed_new_stoika.append(elem)
                    else:
                        if Diametr_dlina_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_dlina_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_stoika.append(elem)

                    setattr(new_stoika, field.split('_')[-1], data[field])

                if 'kreplenie' in field:
                    if existed_new_nadzem_stal_gazoprovod:
                        for gzp in existed_new_stoika:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_stoika.remove(gzp)

                    if Nadzem_stal_gazoprovod.objects.filter(
                        kreplenie=data[field]):
                        for elem in Nadzem_stal_gazoprovod.objects.filter(
                            kreplenie=data[field]):
                            if list(elem.truba.all()) == list(new_nadzem_stal_gazoprovod.truba.all()):
                                existed_new_nadzem_stal_gazoprovod.append(elem)

                    new_nadzem_stal_gazoprovod.kreplenie = data[field]

            elif field in podzem_poliet_gaz_fields:
                if 'poliet_otvod' in field:

                    if 'diametr' in field:

                        if existed_new_otvod:
                            for gzp in existed_new_otvod:
                                if getattr(gzp, 'diametr') != data[field]:
                                    existed_new_otvod.remove(gzp)

                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_otvod.append(elem)
                        setattr(new_otvod, field.split('_')[-1], data[field])

                    else:

                        if existed_new_otvod:
                            for gzp in existed_new_otvod:
                                if getattr(gzp, 'kolvo') != data[field]:
                                    existed_new_otvod.remove(gzp)

                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_otvod.append(elem)

                        setattr(new_otvod, 'kolvo', data[field])

                elif field == 'poliet_troinik1':

                    if existed_new_troinik:
                        for gzp in existed_new_troinik:
                            if getattr(gzp, 'diametr1') != data[field]:
                                existed_new_troinik.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr1=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr1=data[field]):
                            existed_new_troinik.append(elem)

                    new_troinik.diametr1 = data[field]

                elif field == 'poliet_troinik2':

                    if existed_new_troinik:
                        for gzp in existed_new_troinik:
                            if getattr(gzp, 'diametr2') != data[field]:
                                existed_new_troinik.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr2=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr2=data[field]):
                            existed_new_troinik.append(elem)

                    new_troinik.diametr2 = data[field]

                elif field == 'poliet_troinik3':

                    if existed_new_troinik:
                        for gzp in existed_new_troinik:
                            if getattr(gzp, 'diametr3') != data[field]:
                                existed_new_troinik.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr3=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr3=data[field]):
                            existed_new_troinik.append(elem)

                    new_troinik.diametr3 = data[field]

                elif field == 'poliet_troinik_dlina':

                    if existed_new_troinik:
                        for gzp in existed_new_troinik:
                            if getattr(gzp, 'kolvo') != data[field]:
                                existed_new_troinik.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(kolvo=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(kolvo=data[field]):
                            existed_new_troinik.append(elem)

                    new_troinik.kolvo = data[field]

                elif field == 'sedelka_poliet_troinik1':

                    if existed_new_sedelka:
                        for gzp in existed_new_sedelka:
                            if getattr(gzp, 'diametr1') != data[field]:
                                existed_new_sedelka.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr1=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr1=data[field]):
                            existed_new_sedelka.append(elem)

                    new_sedelka.diametr1 = data[field]

                elif field == 'sedelka_poliet_troinik2':

                    if existed_new_sedelka:
                        for gzp in existed_new_sedelka:
                            if getattr(gzp, 'diametr2') != data[field]:
                                existed_new_sedelka.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr2=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr2=data[field]):
                            existed_new_sedelka.append(elem)

                    new_sedelka.diametr2 = data[field]

                elif field == 'sedelka_poliet_troinik3':

                    if existed_new_sedelka:
                        for gzp in existed_new_sedelka:
                            if getattr(gzp, 'diametr3') != data[field]:
                                existed_new_sedelka.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(diametr3=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(diametr3=data[field]):
                            existed_new_sedelka.append(elem)

                    new_sedelka.diametr3 = data[field]

                elif field == 'sedelka_poliet_kolvo':

                    if existed_new_sedelka:
                        for gzp in existed_new_sedelka:
                            if getattr(gzp, 'kolvo') != data[field]:
                                existed_new_sedelka.remove(gzp)

                    if Diametrs_3_kolvo.objects.filter(kolvo=data[field]):
                        for elem in Diametrs_3_kolvo.objects.filter(kolvo=data[field]):
                            existed_new_sedelka.append(elem)

                    new_sedelka.kolvo = data[field]

                elif 'poliet_zaglushka' in field:

                    if field.split('_')[-1] == 'diametr':
                        if existed_new_zaglushka:
                            for gzp in existed_new_zaglushka:
                                if getattr(gzp, 'diametr') != data[field]:
                                    existed_new_zaglushka.remove(gzp)

                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_zaglushka.append(elem)

                        setattr(new_zaglushka, 'diametr', data[field])
                    else:
                        if existed_new_zaglushka:
                            for gzp in existed_new_zaglushka:
                                if getattr(gzp, 'kolvo') != data[field]:
                                    existed_new_zaglushka.remove(gzp)

                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_zaglushka.append(elem)

                        setattr(new_zaglushka, 'kolvo', data[field])

                elif 'poliet_kran_shar' in field:

                    if field.split('_')[-1] == 'diametr':
                        if existed_new_zaglushka:
                            for gzp in existed_new_zaglushka:
                                if getattr(gzp, 'diametr') != data[field]:
                                    existed_new_zaglushka.remove(gzp)

                        if Diametr_kolvo.objects.filter(diametr=data[field]):
                            for elem in Diametr_kolvo.objects.filter(diametr=data[field]):
                                existed_new_kran.append(elem)
                    else:
                        if existed_new_zaglushka:
                            for gzp in existed_new_zaglushka:
                                if getattr(gzp, 'kolvo') != data[field]:
                                    existed_new_zaglushka.remove(gzp)
                        if Diametr_kolvo.objects.filter(kolvo=data[field]):
                            for elem in Diametr_kolvo.objects.filter(kolvo=data[field]):
                                existed_new_kran.append(elem)

                    setattr(new_kran, field.split('_')[-1], data[field])

                elif field == 'poliet_lenta_signal_dlina':

                    if existed_new_podzem_poliet_gazoprovod:
                        for gzp in existed_new_podzem_poliet_gazoprovod:
                            if getattr(gzp, 'lenta') != data[field]:
                                existed_new_podzem_poliet_gazoprovod.remove(gzp)

                    if Podzem_polietilen_gazoprovod.objects.filter(lenta=data[field]):
                        for elem in Podzem_polietilen_gazoprovod.objects.filter(lenta=data[field]):
                            if (list(elem.truba.all()) == list(new_podzem_poliet_gazoprovod.truba.all())
                                and list(elem.mufta.all()) == list(new_podzem_poliet_gazoprovod.mufta.all())):

                                existed_new_podzem_poliet_gazoprovod.append(elem)

                    new_podzem_poliet_gazoprovod.lenta = data[field]

                elif field == 'poliet_opoznavat_znak_kolvo':

                    if existed_new_podzem_poliet_gazoprovod:
                        for gzp in existed_new_podzem_poliet_gazoprovod:
                            if getattr(gzp, 'znak') != data[field]:
                                existed_new_podzem_poliet_gazoprovod.remove(gzp)

                    if Podzem_polietilen_gazoprovod.objects.filter(znak=data[field]):
                        for elem in Podzem_polietilen_gazoprovod.objects.filter(znak=data[field]):
                            if (list(elem.truba.all()) == list(new_podzem_poliet_gazoprovod.truba.all())
                                and list(elem.mufta.all()) == list(new_podzem_poliet_gazoprovod.mufta.all())):

                                existed_new_podzem_poliet_gazoprovod.append(elem)

                    new_podzem_poliet_gazoprovod.znak = data[field]

            elif field in podzem_stal_gaz_fields:
                if field == 'stalnoi_neraz_soed_PE':

                    if existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                        for gzp in existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_neraz_soed_truba_podzem_stal_gazoprovod.remove(gzp)

                    if Neraz_soed_stal_gazoprovod.objects.filter(PE=data[field]):
                        for elem in Neraz_soed_stal_gazoprovod.objects.filter(PE=data[field]):
                            existed_new_neraz_soed_truba_podzem_stal_gazoprovod.append(elem)

                    new_neraz_soed_truba_podzem_stal_gazoprovod.PE = data[field]

                elif field == 'stalnoi_neraz_soed_ST':

                    if existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                        for gzp in existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_neraz_soed_truba_podzem_stal_gazoprovod.remove(gzp)

                    if Neraz_soed_stal_gazoprovod.objects.filter(ST=data[field]):
                        for elem in Neraz_soed_stal_gazoprovod.objects.filter(ST=data[field]):
                            existed_new_neraz_soed_truba_podzem_stal_gazoprovod.append(elem)

                    new_neraz_soed_truba_podzem_stal_gazoprovod.ST = data[field]

                elif field == 'neraz_soed_kolvo':

                    if existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                        for gzp in existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_neraz_soed_truba_podzem_stal_gazoprovod.remove(gzp)

                    if Neraz_soed_stal_gazoprovod.objects.filter(kolvo=data[field]):
                        for elem in Neraz_soed_stal_gazoprovod.objects.filter(kolvo=data[field]):
                            existed_new_neraz_soed_truba_podzem_stal_gazoprovod.append(elem)

                    new_neraz_soed_truba_podzem_stal_gazoprovod.kolvo = data[field]

                elif field == 'stalnoi_kontrolnaya_trubka':

                    if existed_new_podzem_stal_gazoprovod:
                        for gzp in existed_new_podzem_stal_gazoprovod:
                            if getattr(gzp,'stalnoi_kontrolnaya_trubka') != data[field]:
                                existed_new_podzem_stal_gazoprovod.remove(gzp)

                    if Podzem_stal_gazoprovod.objects.filter(kontrolnaya_trubka=data[field]):
                        for elem in Podzem_stal_gazoprovod.objects.filter(kontrolnaya_trubka=data[field]):
                            if list(elem.truba.all()) == list(new_podzem_stal_gazoprovod.truba.all()):
                                existed_new_podzem_stal_gazoprovod.append(elem)

                    new_podzem_stal_gazoprovod.kontrolnaya_trubka = data[field]

                elif field == 'stalnoi_otvod':

                    if existed_new_podzem_stal_gazoprovod:
                        for gzp in existed_new_podzem_stal_gazoprovod:
                            if getattr(gzp, 'otvod_90') != data[field]:
                                existed_new_podzem_stal_gazoprovod.remove(gzp)

                    if Podzem_stal_gazoprovod.objects.filter(otvod_90=data[field]):
                        for elem in Podzem_stal_gazoprovod.objects.filter(otvod_90=data[field]):
                            if list(elem.truba.all()) == list(new_podzem_stal_gazoprovod.truba.all()):
                                existed_new_podzem_stal_gazoprovod.append(elem)

                    new_podzem_stal_gazoprovod.otvod_90 = data[field]

                elif field == 'stal_opoznavat_znak':

                    if existed_new_podzem_stal_gazoprovod:
                        for gzp in existed_new_podzem_stal_gazoprovod:
                            if getattr(gzp, 'opoznavat_znak') != data[field]:
                                existed_new_podzem_stal_gazoprovod.remove(gzp)

                    if Podzem_stal_gazoprovod.objects.filter(opoznavat_znak=data[field]):
                        for elem in Podzem_stal_gazoprovod.objects.filter(opoznavat_znak=data[field]):
                            if list(elem.truba.all()) == list(new_podzem_stal_gazoprovod.truba.all()):
                                existed_new_podzem_stal_gazoprovod.append(elem)

                    new_podzem_stal_gazoprovod.opoznavat_znak = data[field]

            elif field in ks11_fields:

                if 'podryadchik' in field:

                    if existed_person_podryadchik:
                        for gzp in existed_person_podryadchik:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_podryadchik.remove(gzp)

                    if 'fio' in field:
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_podryadchik.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_podryadchik.append(elem)

                    setattr(person_podryadchik, field.split('_')[-1], data[field])


                if 'predsedatel' in field:

                    if existed_person_preds:
                        for gzp in existed_person_preds:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_preds.remove(gzp)

                    if 'fio' in field:
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_preds.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_preds.append(elem)

                    setattr(person_preds, field.split('_')[-1], data[field])


                if 'predstav_ekspl' in field:

                    if existed_person_predstav_eks:
                        for gzp in existed_person_predstav_eks:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_predstav_eks.remove(gzp)

                    if 'fio' in field:
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_predstav_eks.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_predstav_eks.append(elem)

                    setattr(person_predstav_eks, field.split('_')[-1], data[field])


                if 'predstav_proekt' in field:

                    if existed_person_predstav_proek:
                        for gzp in existed_person_predstav_proek:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_predstav_proek.remove(gzp)

                    if 'fio' in field:
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_predstav_proek.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_predstav_proek.append(elem)

                    setattr(person_predstav_proek, field.split('_')[-1], data[field])

            elif field in kontragent_fields:

                if 'podryadchik' in field:

                    if existed_new_podpisant:
                        for gzp in existed_new_podpisant:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_podpisant.remove(gzp)

                    if 'fio' in field:
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_new_podpisant.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_new_podpisant.append(elem)


                elif field == 'podpisant':

                    if existed_new_podpisant:
                        for gzp in existed_new_podpisant:
                            if getattr(gzp, 'fio') != data[field]:
                                existed_new_podpisant.remove(gzp)

                    if Person.objects.filter(fio=data[field]):
                        for elem in Person.objects.filter(fio=data[field]):
                            existed_new_podpisant.append(elem)

                    setattr(new_podpisant, field.split('_')[-1], data[field])

                else:

                    if existed_new_kontragent:
                        for gzp in existed_new_kontragent:
                            if getattr(gzp, field) != data[field]:
                                existed_new_kontragent.remove(gzp)

                    if 'name' in field:
                        if Kontragent.objects.filter(name_kontragent=data[field]):
                            for elem in Kontragent.objects.filter(name_kontragent=data[field]):
                                existed_new_kontragent.append(elem)
                    elif 'INN' in field:
                        if Kontragent.objects.filter(INN=data[field]):
                            for elem in Kontragent.objects.filter(INN=data[field]):
                                existed_new_kontragent.append(elem)
                    elif 'KPP' in field:
                        if Kontragent.objects.filter(KPP=data[field]):
                            for elem in Kontragent.objects.filter(KPP=data[field]):
                                existed_new_kontragent.append(elem)
                    elif 'Ur_address' in field:
                        if Kontragent.objects.filter(Ur_address=data[field]):
                            for elem in Kontragent.objects.filter(Ur_address=data[field]):
                                existed_new_kontragent.append(elem)
                    else:
                        if Kontragent.objects.filter(telephone=data[field]):
                            for elem in Kontragent.objects.filter(telephone=data[field]):
                                existed_new_kontragent.append(elem)

                    setattr(new_kontragent, field, data[field])

            elif field in ks2_fields:

                if 'ks2_zakazchik' in field:

                    if existed_person_zakazchik:
                        for gzp in existed_person_zakazchik:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_zakazchik.remove(gzp)

                    if field.split('_')[-1] == 'fio':
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_zakazchik.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_zakazchik.append(elem)

                    setattr(person_zakazchik, field.split('_')[-1], data[field])

                if 'ks2_podryadchik' in field:

                    if existed_person_ks2_podryadchik:
                        for gzp in existed_person_ks2_podryadchik:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_person_ks2_podryadchik.remove(gzp)

                    if field.split('_')[-1] == 'fio':
                        if Person.objects.filter(fio=data[field]):
                            for elem in Person.objects.filter(fio=data[field]):
                                existed_person_ks2_podryadchik.append(elem)
                    else:
                        if Person.objects.filter(post=data[field]):
                            for elem in Person.objects.filter(post=data[field]):
                                existed_person_ks2_podryadchik.append(elem)

                    setattr(person_ks2_podryadchik, field.split('_')[-1], data[field])

                if field == 'object':
                    setattr(object, 'ks2_object', data[field])
                if field == 'stroika':
                    setattr(object, 'ks2_stroika', data[field])

            elif field in svarshik1_fields:

                if 'date' in field:

                    if existed_new_svarshik1:
                        for gzp in existed_new_svarshik1:
                            if getattr(gzp, 'date_svarki') != datetime.datetime.strptime(data[field],"%Y-%m-%d"):
                                existed_new_svarshik1.remove(gzp)

                    if Svarshik.objects.filter(date_svarki=datetime.datetime.strptime(data[field],"%Y-%m-%d"), type='Полиэтилен'):
                        for elem in Svarshik.objects.filter(date_svarki=datetime.datetime.strptime(data[field],"%Y-%m-%d"), type='Полиэтилен'):
                            existed_new_svarshik1.append(elem)

                    setattr(new_svarshik1, 'date_svarki', datetime.datetime.strptime(data[field],"%Y-%m-%d"))

                else:

                    if existed_new_svarshik1:
                        for gzp in existed_new_svarshik1:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_svarshik1.remove(gzp)

                    if 'fio' in field.split('_')[-1]:
                        if Svarshik.objects.filter(fio=data[field], type='Полиэтилен'):
                            for elem in Svarshik.objects.filter(fio=data[field], type='Полиэтилен'):
                                existed_new_svarshik1.append(elem)
                    if 'post' in field.split('_')[-1]:
                        if Svarshik.objects.filter(post=data[field], type='Полиэтилен'):
                            for elem in Svarshik.objects.filter(post=data[field], type='Полиэтилен'):
                                existed_new_svarshik1.append(elem)
                    if 'diametr' in field.split('_')[-1]:
                        if Svarshik.objects.filter(diametr=data[field], type='Полиэтилен'):
                            for elem in Svarshik.objects.filter(diametr=data[field], type='Полиэтилен'):
                                existed_new_svarshik1.append(elem)
                    if 'kolvo' in field.split('_')[-1]:
                        if Svarshik.objects.filter(kolvo=data[field], type='Полиэтилен'):
                            for elem in Svarshik.objects.filter(kolvo=data[field], type='Полиэтилен'):
                                existed_new_svarshik1.append(elem)

                    setattr(new_svarshik1, field.split('_')[-1], data[field])

                setattr(new_svarshik1, 'type', 'Полиэтилен')

            elif field in svarshik2_fields:

                if 'date' in field:

                    if existed_new_svarshik2:
                        for gzp in existed_new_svarshik2:
                            if getattr(gzp, 'date_svarki') != datetime.datetime.strptime(data[field],"%Y-%m-%d"):
                                existed_new_svarshik2.remove(gzp)

                    if Svarshik.objects.filter(date_svarki=datetime.datetime.strptime(data[field],"%Y-%m-%d"), type='Сталь'):
                        for elem in Svarshik.objects.filter(date_svarki=datetime.datetime.strptime(data[field],"%Y-%m-%d"), type='Сталь'):
                            existed_new_svarshik2.append(elem)

                    setattr(new_svarshik2, 'date_svarki', datetime.datetime.strptime(data[field],"%Y-%m-%d"))

                else:

                    if existed_new_svarshik2:
                        for gzp in existed_new_svarshik2:
                            if getattr(gzp, field.split('_')[-1]) != data[field]:
                                existed_new_svarshik2.remove(gzp)

                    if 'fio' in field.split('_')[-1]:
                        if Svarshik.objects.filter(fio=data[field], type='Сталь'):
                            for elem in Svarshik.objects.filter(fio=data[field], type='Сталь'):
                                existed_new_svarshik2.append(elem)
                    if 'post' in field.split('_')[-1]:
                        if Svarshik.objects.filter(post=data[field], type='Сталь'):
                            for elem in Svarshik.objects.filter(post=data[field], type='Сталь'):
                                existed_new_svarshik2.append(elem)
                    if 'diametr' in field.split('_')[-1]:
                        if Svarshik.objects.filter(diametr=data[field], type='Сталь'):
                            for elem in Svarshik.objects.filter(diametr=data[field], type='Сталь'):
                                existed_new_svarshik2.append(elem)
                    if 'kolvo' in field.split('_')[-1]:
                        if Svarshik.objects.filter(kolvo=data[field], type='Сталь'):
                            for elem in Svarshik.objects.filter(kolvo=data[field], type='Сталь'):
                                existed_new_svarshik2.append(elem)
                    setattr(new_svarshik2, field.split('_')[-1], data[field])

                setattr(new_svarshik2, 'type', 'Сталь')

            elif field in zashit_futlyar_fields:

                if existed_new_zashit_futlyar:
                    for gzp in existed_new_zashit_futlyar:
                        if getattr(gzp, field.split('_')[-1]) != data[field]:
                            existed_new_zashit_futlyar.remove(gzp)

                if 'diametr' in field.split('_')[-1]:
                    if Diametr_x_dlina.objects.filter(diametr=data[field]):
                        for elem in Diametr_x_dlina.objects.filter(diametr=data[field]):
                            existed_new_zashit_futlyar.append(elem)
                elif 'x' in field.split('_')[-1]:
                    if Diametr_x_dlina.objects.filter(x=data[field]):
                        for elem in Diametr_x_dlina.objects.filter(x=data[field]):
                            existed_new_zashit_futlyar.append(elem)
                else:
                    if Diametr_x_dlina.objects.filter(dlina=data[field]):
                        for elem in Diametr_x_dlina.objects.filter(dlina=data[field]):
                            existed_new_zashit_futlyar.append(elem)

                setattr(new_zashit_futlyar, field.split('_')[-1], data[field])

            elif field in opora_pod_gazoprovod_fields:
                if existed_new_opora:
                    for gzp in existed_new_opora:
                        if getattr(gzp, field.split('_')[-1]) != data[field]:
                            existed_new_opora.remove(gzp)

                if 'diametr' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(diametr=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(diametr=data[field]):
                            existed_new_opora.append(elem)
                elif 'x' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(x=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(x=data[field]):
                            existed_new_opora.append(elem)
                elif 'dlina' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(dlina=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(dlina=data[field]):
                            existed_new_opora.append(elem)
                else:
                    if Diametr_x_dlina_kolvo.objects.filter(kolvo=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(kolvo=data[field]):
                            existed_new_opora.append(elem)

                setattr(new_opora, field.split('_')[-1], data[field])

            elif field in futlyar_na_vyhode_fields:

                if existed_new_futlyar_na_vyh:
                    for gzp in existed_new_futlyar_na_vyh:
                        if getattr(gzp, field.split('_')[-1]) != data[field]:
                            existed_new_futlyar_na_vyh.remove(gzp)

                if 'diametr' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(diametr=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(diametr=data[field]):
                            existed_new_futlyar_na_vyh.append(elem)
                elif 'x' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(x=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(x=data[field]):
                            existed_new_futlyar_na_vyh.append(elem)
                elif 'dlina' in field.split('_')[-1]:
                    if Diametr_x_dlina_kolvo.objects.filter(dlina=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(dlina=data[field]):
                            existed_new_futlyar_na_vyh.append(elem)
                else:
                    if Diametr_x_dlina_kolvo.objects.filter(kolvo=data[field]):
                        for elem in Diametr_x_dlina_kolvo.objects.filter(kolvo=data[field]):
                            existed_new_futlyar_na_vyh.append(elem)

                setattr(new_futlyar_na_vyh, field.split('_')[-1], data[field])

            elif field in smeta_fields:

                if 'date' in field:

                    if field != 'date_nach_rabot':

                        if existed_new_smeta:
                            for gzp in existed_new_smeta:
                                if getattr(gzp, field) != datetime.datetime.strptime(data[field],"%Y-%m-%d"):
                                    existed_new_smeta.remove(gzp)
                        if 'ks2' in field:
                            if Smeta.objects.filter(date_ks2=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                for elem in Smeta.objects.filter(date_ks2=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                    existed_new_smeta.append(elem)
                        elif 'nach_zakr' in field:
                            if Smeta.objects.filter(date_nach_zakr=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                for elem in Smeta.objects.filter(date_nach_zakr=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                    existed_new_smeta.append(elem)
                        elif 'kon_zakr' in field:
                            if Smeta.objects.filter(date_kon_zakr=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                for elem in Smeta.objects.filter(date_kon_zakr=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                    existed_new_smeta.append(elem)
                        else:
                            if Smeta.objects.filter(date_dogovor=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                for elem in Smeta.objects.filter(date_dogovor=datetime.datetime.strptime(data[field],"%Y-%m-%d")):
                                    existed_new_smeta.append(elem)

                        setattr(new_smeta, field, datetime.datetime.strptime(data[field],"%Y-%m-%d"))
                    else:
                        date_nach_rabot = field
                        full_date_nach_rabot = f"{data[field]}-01"

                        if existed_new_smeta:
                            for gzp in existed_new_smeta:
                                if getattr(gzp, field) != datetime.datetime.strptime(full_date_nach_rabot,"%Y-%m-%d"):
                                    existed_new_smeta.remove(gzp)

                        if Smeta.objects.filter(date_nach_rabot=datetime.datetime.strptime(full_date_nach_rabot,"%Y-%m-%d")):
                            for elem in Smeta.objects.filter(date_nach_rabot=datetime.datetime.strptime(full_date_nach_rabot,"%Y-%m-%d")):
                                existed_new_smeta.append(elem)

                        new_smeta.date_nach_rabot = full_date_nach_rabot


                else:
                    if existed_new_smeta:
                        for gzp in existed_new_smeta:
                            if getattr(gzp, field) != data[field]:
                                existed_new_smeta.remove(gzp)

                    if 'proektnoi_smeti' in field:
                        if Smeta.objects.filter(summa_proektnoi_smeti=data[field]):
                            for elem in Smeta.objects.filter(summa_proektnoi_smeti=data[field]):
                                existed_new_smeta.append(elem)
                    elif 'utv_smeti' in field:
                        if Smeta.objects.filter(summa_utv_smeti=data[field]):
                            for elem in Smeta.objects.filter(summa_utv_smeti=data[field]):
                                existed_new_smeta.append(elem)
                    elif 'ks2_bez_nds' in field:
                        if Smeta.objects.filter(summa_ks2_bez_nds=data[field]):
                            for elem in Smeta.objects.filter(summa_ks2_bez_nds=data[field]):
                                existed_new_smeta.append(elem)
                    else:
                        if Smeta.objects.filter(nomer_dogovor=data[field]):
                            for elem in Smeta.objects.filter(nomer_dogovor=data[field]):
                                existed_new_smeta.append(elem)

                    setattr(new_smeta, field, data[field])

            elif field == 'district':
                if not new_district:
                    new_district = District.objects.get_or_create(name=data[field])[0]
                    new_district.save()
                    object.district = new_district


            elif field == 'tehnadzor_fio' or field == 'tehnadzor_post':

                if field == 'tehnadzor_fio':
                    new_tehnadzor_person.fio = data[field]


                    if flag_teh:
                        if Person.objects.filter(fio=new_tehnadzor_person.fio,
                            post=new_tehnadzor_person.post):
                            new_tehnadzor_person = Person.objects.filter(fio=new_tehnadzor_person.fio,
                                post=new_tehnadzor_person.post)[0]
                        else:
                            new_tehnadzor_person.save()


                        if 'district' in data.keys():

                            if not new_district:
                                new_district = District.objects.get_or_create(name=data[field])[0]
                                new_district.save()
                                object.district = new_district

                            if TehNadzor.objects.filter(person=new_tehnadzor_person, district=new_district):
                                new_tehnadzor = TehNadzor.objects.filter(person=new_tehnadzor_person, district=new_district)[0]
                            else:
                                new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district=new_district)[0]
                            new_tehnadzor.save()
                            object.tehnadzor = new_tehnadzor
                            object.tehnadzor.save()

                        else:
                            if object.district:
                                if TehNadzor.objects.filter(person=new_tehnadzor_person, district=object.district):
                                    new_tehnadzor = TehNadzor.objects.filter(person=new_tehnadzor_person, district=object.district)[0]
                                else:
                                    new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district=object.district)[0]
                            else:
                                new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district='')[0]
                            new_tehnadzor.save()
                            object.tehnadzor = new_tehnadzor
                            object.tehnadzor.save()

                    else:
                        flag_teh = True

                if field == 'tehnadzor_post':

                    new_tehnadzor_person.post = data[field]

                    if flag_teh:

                        if Person.objects.filter(fio=new_tehnadzor_person.fio,
                            post=new_tehnadzor_person.post):
                            new_tehnadzor_person = Person.objects.filter(fio=new_tehnadzor_person.fio,
                                post=new_tehnadzor_person.post)[0]
                        else:
                            new_tehnadzor_person.save()

                        if 'district' in data.keys():

                            if not new_district:
                                new_district = District.objects.get_or_create(name=data[field])[0]
                                new_district.save()
                                object.district = new_district

                            if TehNadzor.objects.filter(person=new_tehnadzor_person, district=new_district):
                                new_tehnadzor = TehNadzor.objects.filter(person=new_tehnadzor_person, district=new_district)[0]
                            else:
                                new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district=new_district)[0]
                            new_tehnadzor.save()
                            object.tehnadzor = new_tehnadzor
                            object.tehnadzor.save()

                        else:
                            if object.district:
                                if TehNadzor.objects.filter(person=new_tehnadzor_person, district=object.district):
                                    new_tehnadzor = TehNadzor.objects.filter(person=new_tehnadzor_person, district=object.district)[0]
                                else:
                                    new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district=object.district)[0]
                            else:
                                new_tehnadzor = TehNadzor.objects.get_or_create(person=new_tehnadzor_person, district='')[0]
                            new_tehnadzor.save()
                            object.tehnadzor = new_tehnadzor
                            object.tehnadzor.save()

                    else:
                        flag_teh = True

            else:
                if field in fields.keys():
                    if field not in ('gazoprovod', 'prodavlivanie'):
                        try:
                            setattr(object, field, data[field])
                        except:
                            pass

    if existed_new_gip:
        object.gip = existed_new_gip[0]
    else:
        if new_gip:
            new_gip.save()
            object.gip = new_gip

    if existed_new_nadzem_stal_gazoprovod_izolir:
        new_nadzem_stal_gazoprovod.izolir_soed = existed_new_nadzem_stal_gazoprovod_izolir[0]
    else:
        if new_nadzem_stal_gazoprovod_izolir:
            new_nadzem_stal_gazoprovod_izolir.save()
            new_nadzem_stal_gazoprovod.izolir_soed = new_nadzem_stal_gazoprovod_izolir

    if existed_new_nadzem_stal_gazoprovod_kran:
        new_nadzem_stal_gazoprovod.kran_stal = existed_new_nadzem_stal_gazoprovod_kran[0]
    else:
        if new_nadzem_stal_gazoprovod_kran:
            new_nadzem_stal_gazoprovod_kran.save()
            new_nadzem_stal_gazoprovod.kran_stal = new_nadzem_stal_gazoprovod_kran

    if existed_new_nadzem_stal_gazoprovod_otvod:
        new_nadzem_stal_gazoprovod.otvod = existed_new_nadzem_stal_gazoprovod_otvod[0]
    else:
        if new_nadzem_stal_gazoprovod_otvod:
            new_nadzem_stal_gazoprovod_otvod.save()
            new_nadzem_stal_gazoprovod.otvod = new_nadzem_stal_gazoprovod_otvod

    if existed_new_nadzem_stal_gazoprovod_cokol:
        new_nadzem_stal_gazoprovod.cokolnyi_vvod = existed_new_nadzem_stal_gazoprovod_cokol[0]
    else:
        if new_nadzem_stal_gazoprovod_cokol:
            new_nadzem_stal_gazoprovod_cokol.save()
            new_nadzem_stal_gazoprovod.cokolnyi_vvod = new_nadzem_stal_gazoprovod_cokol

    if existed_new_stoika:
        new_nadzem_stal_gazoprovod.stoika = existed_new_stoika[0]
    else:
        if new_stoika:
            new_stoika.save()
            new_nadzem_stal_gazoprovod.stoika = new_stoika

    if new_nadzem_stal_gazoprovod:
        new_nadzem_stal_gazoprovod.save()
        object.gazoprovod_nadzem_stal = new_nadzem_stal_gazoprovod

    if existed_new_otvod:
        new_podzem_poliet_gazoprovod.otvod = existed_new_otvod[0]
    else:
        if new_otvod:
            new_otvod.save()
            new_podzem_poliet_gazoprovod.otvod = new_otvod

    if existed_new_troinik:
        new_podzem_poliet_gazoprovod.troinik = existed_new_troinik[0]
    else:
        if new_troinik:
            new_troinik.save()
            new_podzem_poliet_gazoprovod.troinik = new_troinik

    if existed_new_sedelka:
        new_podzem_poliet_gazoprovod.sedelka = existed_new_sedelka[0]
    else:
        if new_sedelka:
            new_sedelka.save()
            new_podzem_poliet_gazoprovod.sedelka = new_sedelka

    if existed_new_zaglushka:
        new_podzem_poliet_gazoprovod.zaglushka = existed_new_zaglushka[0]
    else:
        if new_zaglushka:
            new_zaglushka.save()
            new_podzem_poliet_gazoprovod.zaglushka = new_zaglushka

    if existed_new_kran:
        new_podzem_poliet_gazoprovod.kran = existed_new_kran[0]
    else:
        if new_kran:
            new_kran.save()
            new_podzem_poliet_gazoprovod.kran = new_kran

    if new_podzem_poliet_gazoprovod:
        new_podzem_poliet_gazoprovod.save()
        object.gazoprovod_podzem_poliet = new_podzem_poliet_gazoprovod


    if existed_new_neraz_soed_truba_podzem_stal_gazoprovod:
        new_podzem_stal_gazoprovod.neraz_soed = existed_new_neraz_soed_truba_podzem_stal_gazoprovod[0]
    else:
        if new_neraz_soed_truba_podzem_stal_gazoprovod:
            new_neraz_soed_truba_podzem_stal_gazoprovod.save()
            new_podzem_stal_gazoprovod.neraz_soed = new_neraz_soed_truba_podzem_stal_gazoprovod

    if new_podzem_stal_gazoprovod:
        new_podzem_stal_gazoprovod.save()
        object.gazoprovod_podzem_stal = new_podzem_stal_gazoprovod

    if existed_new_kontragent:
        object.kontragent = existed_new_kontragent[0]
    else:
        if new_kontragent:
            new_kontragent.save()
            object.kontragent = new_kontragent

    if existed_new_podpisant:
        object.kontragent.podpisant = existed_new_podpisant[0]
    else:
        if new_podpisant:
            new_podpisant.save()
            object.kontragent.podpisant = new_podpisant

    if existed_person_podryadchik:
        object.kontragent.podpisant = existed_person_podryadchik[0]
    else:
        if person_podryadchik:
            person_podryadchik.save()
            object.kontragent.podpisant = person_podryadchik

    if existed_person_preds:
        object.ks11_predsedatel = existed_person_preds[0]
    else:
        if person_preds:
            person_preds.save()
            object.ks11_predsedatel = person_preds

    if existed_person_predstav_eks:
        object.ks11_predstav_ekspl = existed_person_predstav_eks[0]
    else:
        if person_predstav_eks:
            person_predstav_eks.save()
            object.ks11_predstav_ekspl = person_predstav_eks

    if existed_person_predstav_proek:
        object.ks11_predstav_proekt = existed_person_predstav_proek[0]
    else:
        if person_predstav_proek:
            person_predstav_proek.save()
            object.ks11_predstav_proekt = person_predstav_proek

    if existed_person_zakazchik:
        object.ks2_zakazchik = existed_person_zakazchik[0]
    else:
        if person_zakazchik:
            person_zakazchik.save()
            object.ks2_zakazchik = person_zakazchik

    if existed_person_ks2_podryadchik:
        object.ks2_podryadchik = existed_person_ks2_podryadchik[0]
    else:
        if person_ks2_podryadchik:
            person_ks2_podryadchik.save()
            object.ks2_podryadchik = person_ks2_podryadchik

    if existed_new_rostehnadzor:
        if existed_new_rostehnadzor_truba:
            a = Rostehnadzor.objects.get_or_create(
                    pk1=existed_new_rostehnadzor[0].pk1,
                    pk1_diam=existed_new_rostehnadzor[0].pk1_diam,
                    pk2=existed_new_rostehnadzor[0].pk2,
                    pk2_diam=existed_new_rostehnadzor[0].pk2_diam,
                    truba=existed_new_rostehnadzor_truba[0])[0]
            a.save()
            object.rostehnadzor = a
        else:
            if new_rostehnadzor_truba:
                new_rostehnadzor_truba.save()
                a = Rostehnadzor.objects.get_or_create(
                        pk1=existed_new_rostehnadzor[0].pk1,
                        pk1_diam=existed_new_rostehnadzor[0].pk1_diam,
                        pk2=existed_new_rostehnadzor[0].pk2,
                        pk2_diam=existed_new_rostehnadzor[0].pk2_diam,
                        truba=new_rostehnadzor_truba)[0]
                a.save()
                object.rostehnadzor = a
            else:
                object.rostehnadzor = existed_new_rostehnadzor

    else:
        if new_rostehnadzor:
            if existed_new_rostehnadzor_truba:
                b = Rostehnadzor.objects.get_or_create(
                        pk1=new_rostehnadzor.pk1,
                        pk1_diam=new_rostehnadzor.pk1_diam,
                        pk2=new_rostehnadzor.pk2,
                        pk2_diam=new_rostehnadzor.pk2_diam,
                        truba=existed_new_rostehnadzor_truba[0])[0]
                b.save()
                object.rostehnadzor = b
            else:
                if new_rostehnadzor_truba:
                    new_rostehnadzor_truba.save()
                    b = Rostehnadzor.objects.get_or_create(
                            pk1=new_rostehnadzor.pk1,
                            pk1_diam=new_rostehnadzor.pk1_diam,
                            pk2=new_rostehnadzor.pk2,
                            pk2_diam=new_rostehnadzor.pk2_diam,
                            truba=new_rostehnadzor_truba)[0]
                    b.save()
                    object.rostehnadzor = b
                else:
                    object.rostehnadzor = new_rostehnadzor

        else:
            object.rostehnadzor = None


    if existed_new_prodavl:
        if existed_new_prodavl_truba:
            a = Prodavlivanie.objects.get_or_create(
                    pk0_1_diam=existed_new_prodavl[0].pk0_1_diam,
                    pk0_2_diam=existed_new_prodavl[0].pk0_2_diam,
                    truba=existed_new_prodavl_truba[0])[0]
            a.save()
            object.prodavl = a
        else:
            if new_prodavl_truba:
                new_prodavl_truba.save()
                a = Prodavlivanie.objects.get_or_create(
                    pk0_1_diam=existed_new_prodavl[0].pk0_1_diam,
                    pk0_2_diam=existed_new_prodavl[0].pk0_2_diam,
                    truba=new_prodavl_truba)[0]
                a.save()
                object.prodavl = a
            else:
                object.prodavl = existed_new_prodavl

    else:
        if new_prodavl:
            if existed_new_prodavl_truba:
                b = Prodavlivanie.objects.get_or_create(
                    pk0_1_diam=new_prodavl.pk0_1_diam,
                    pk0_2_diam=new_prodavl.pk0_2_diam,
                    truba=existed_new_prodavl_truba[0])[0]
                b.save()
                object.prodavl = b
            else:
                if new_prodavl_truba:
                    new_prodavl_truba.save()
                    b = Prodavlivanie.objects.get_or_create(
                        pk0_1_diam=new_prodavl.pk0_1_diam,
                        pk0_2_diam=new_prodavl.pk0_2_diam,
                        truba=new_prodavl_truba)[0]
                    b.save()
                    object.prodavl = b
                else:
                    object.prodavl = new_prodavl

        else:
            object.prodavl = None

    if existed_new_svarshik1:
        object.svarshik1 = existed_new_svarshik1[0]
    else:
        if new_svarshik1:
            new_svarshik1.save()
            object.svarshik1 = new_svarshik1

    if existed_new_svarshik2:
        object.svarshik1 = existed_new_svarshik2[0]
    else:
        if new_svarshik2:
            new_svarshik2.save()
            object.svarshik1 = new_svarshik2

    if existed_new_zashit_futlyar:
        object.zashitnyi_futlyar = existed_new_zashit_futlyar[0]
    else:
        if new_zashit_futlyar:
            new_zashit_futlyar.save()
            object.zashitnyi_futlyar = new_zashit_futlyar

    if existed_new_opora:
        object.opora = existed_new_opora[0]
    else:
        if new_opora:
            new_opora.save()
            object.opora = new_opora

    if existed_new_futlyar_na_vyh:
        object.futlyar_na_vyhode = existed_new_futlyar_na_vyh[0]
    else:
        if new_futlyar_na_vyh:
            new_futlyar_na_vyh.save()
            object.futlyar_na_vyhode = new_futlyar_na_vyh

    completed_smeta = ''
    if existed_new_smeta:
        for ex_smeta in existed_new_smeta:
            if ex_smeta.summa_proektnoi_smeti == new_smeta.summa_proektnoi_smeti and \
                ex_smeta.summa_utv_smeti == new_smeta.summa_utv_smeti and \
                ex_smeta.summa_ks2_bez_nds == new_smeta.summa_ks2_bez_nds and \
                ex_smeta.date_ks2 == new_smeta.date_ks2 and \
                ex_smeta.nomer_dogovor == new_smeta.nomer_dogovor and \
                ex_smeta.date_nach_zakr == new_smeta.date_nach_zakr and \
                ex_smeta.date_kon_zakr == new_smeta.date_kon_zakr and \
                ex_smeta.date_dogovor == new_smeta.date_dogovor and \
                ex_smeta.date_nach_rabot == f"{new_smeta.date_nach_rabot}-01":

                object.smeta = ex_smeta
                completed_smeta = ex_smeta
                break

    if not completed_smeta:
        if new_smeta:
            new_smeta.save()
            object.smeta = new_smeta

    object.save()

    context = {}
    objects = Object.objects.all().order_by('name_object')
    context['objects'] = objects
    context['fields'] = fields
    return render(request, 'object_view.html', {'context':context, 'success':'Данные объекта успешно изменены!'})


def fill_all_lists(context, gips_orgs, gips_list, raions_list, object=None):
    for gip in GIP.objects.all().order_by('fio'):
        gips_list.append(gip.fio)
        gips_orgs[gip.fio]= gip.organization

    for raion in District.objects.all().order_by('name'):
        raions_list.append(raion.name)

    context['raions'] = raions_list
    context['gips_orgs'] = json.dumps(gips_orgs)
    context['gips_list'] = gips_list
    if object:
        context['certificates'] = Certificate.objects.exclude(id__in=object.certificates.values('id'))
    else:
        context['certificates'] = Certificate.objects.all().order_by('name')
    context['tehnadzors'] = TehNadzor.objects.all()
    context['commissions'] = Commission.objects.all()
    context['kontragents'] = Kontragent.objects.all()


def get_indexes_from_list(list):
    count = len(list)
    i = 0
    new_list = []
    while i != count:
        new_list.append(str(i))
        i += 1
    return new_list


def check_all_dates(object):
    if object.Data_razm:
        date_update(object,'Data_razm')
    if object.Date_proekt:
        date_update(object,'Date_proekt')
    if object.Data_zamera1:
        date_update(object,'Data_zamera1')
    if object.Data_zamera2:
        date_update(object,'Data_zamera2')
    if object.Data_sost_project:
        date_update(object,'Data_sost_project')
    if object.Data_razbiv:
        date_update(object,'Data_razbiv')
    if object.Data_produv:
        date_update(object,'Data_produv')
    if object.Data_ukl:
        date_update(object,'Data_ukl')


def date_update(object, date_field):
    setattr(object,date_field,datetime.date.strftime(getattr(object,date_field),"%Y-%m-%d"))


def _add_dann_in_gazopr_v_zashit(data, object, index=''):
    truba1 = Diametr_x.objects.get_or_create(diametr=data['dop_dann_gazopr_v_zashit_diametr'+index],\
                                            x = data['dop_dann_gazopr_v_zashit_kolvo'+index] )[0]
    truba1.save()
    truba2 = Diametr_x.objects.get_or_create(diametr=data['dop_dann_gazopr_v_zashit_truba_diametr'+index],\
                                            x = data['dop_dann_gazopr_v_zashit_truba_diametr'+index] )[0]
    truba1.save()

    object.truba1 = truba1
    object.truba2 = truba2


def _add_dann_in_shar_kran(data, object, index=''):
    diam1 = Diametr.objects.get_or_create(diametr=data['dop_dann_sharovyi_kran'+index])[0]
    diam1.save()
    diam2 = Diametr.objects.get_or_create(diametr=data['dop_dann_sharovyi_kran_mufta'+index])[0]
    diam2.save()
    object.kran = diam1
    object.mufta = diam2


def _add_truba_in_podzem_stal_gazoprovod(data, object, index=''):
    truba_podzem_stal_gazoprovod = Diametr_x_dlina_prim.objects.get_or_create(diametr = data['stalnoi_truba_diametr' + index],\
                                        x = data['stalnoi_truba_kolvo'+ index],\
                                        dlina = data['stalnoi_truba_dlina'+ index],\
                                        prim = data['stalnoi_truba_prim'+ index])[0]

    truba_podzem_stal_gazoprovod.save()
    object.truba.add(truba_podzem_stal_gazoprovod)


def _add_poliet_truba_mufta_in_podzem_poliet_gazoprovod(data, object, index=''):
    truba_podzem_poliet_gazoprovod = Diametr_x_dlina.objects.get_or_create(diametr = data['poliet_truba_diametr' + index],\
                                        x = data['poliet_truba_kolvo'+ index],\
                                        dlina = data['poliet_truba_dlina'+ index])[0]
    truba_podzem_poliet_gazoprovod.save()

    object.truba.add(truba_podzem_poliet_gazoprovod)
    mufta_podzem_poliet_gazoprovod = Diametr_x_dlina.objects.get_or_create(diametr = data['poliet_mufta_diametr' + index],\
                                        x = data['poliet_mufta_kolvo'+ index],\
                                        dlina = data['poliet_mufta_dlina'+ index])[0]
    mufta_podzem_poliet_gazoprovod.save()
    object.mufta.add(mufta_podzem_poliet_gazoprovod)


def _add_truba_in_nadzem_stal_gazoprovod(data, object, index=''):
    truba_nadzem_stal_gazoprovod = Diametr_x_dlina.objects.get_or_create(diametr = data['nadzem_stal_truba_diametr' + index],\
                                        x = data['nadzem_stal_truba_x' + index],\
                                        dlina = data['nadzem_stal_truba_dlina'+ index])[0]
    truba_nadzem_stal_gazoprovod.save()
    object.truba.add(truba_nadzem_stal_gazoprovod)
