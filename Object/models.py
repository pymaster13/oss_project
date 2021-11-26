import datetime

from django.db import models

from .validators import validate_file_extension


class Document(models.Model):
    """Документ"""
    file = models.FileField(upload_to="documents", validators=[validate_file_extension])

class Smeta(models.Model):
    """Данные для сметы"""

    summa_proektnoi_smeti = models.CharField(max_length = 256, verbose_name='Сумма проектной сметы', blank=True, null=True)
    summa_utv_smeti = models.CharField(max_length = 256, verbose_name='Сумма утвержденной сметы', blank=True, null=True)
    summa_ks2_bez_nds = models.CharField(max_length = 256, verbose_name='Сумма КС 2 без НДС', blank=True, null=True)
    date_ks2 = models.DateField(verbose_name='Дата КС 2', blank=True, null=True)
    date_nach_rabot = models.DateField(verbose_name='Дата начала работ', blank=True, null=True)
    date_nach_zakr = models.DateField(verbose_name='Дата начала закрытия', blank=True, null=True)
    date_kon_zakr = models.DateField(verbose_name='Дата окончания закрытия', blank=True, null=True)
    nomer_dogovor = models.CharField(max_length = 256, verbose_name='Номер договора', blank=True, null=True)
    date_dogovor = models.DateField(verbose_name='Дата окончания договора', blank=True, null=True)

    @property
    def date_ks2_html(self):
        if self.date_ks2:
            return datetime.datetime.strftime(self.date_ks2, "%Y-%m-%d")
        else:
            return ''

    @property
    def date_nach_rabot_html(self):
        if self.date_nach_rabot:
            return datetime.datetime.strftime(self.date_nach_rabot, "%Y-%m")
        else:
            return ''

    @property
    def date_nach_zakr_html(self):
        if self.date_nach_zakr:
            return datetime.datetime.strftime(self.date_nach_zakr, "%Y-%m-%d")
        else:
            return ''

    @property
    def date_kon_zakr_html(self):
        if self.date_kon_zakr:
            return datetime.datetime.strftime(self.date_kon_zakr, "%Y-%m-%d")
        else:
            return ''

    @property
    def date_dogovor_html(self):
        if self.date_dogovor:
            return datetime.datetime.strftime(self.date_dogovor, "%Y-%m-%d")
        else:
            return ''

    class Meta:
        verbose_name = 'Смета'
        verbose_name_plural = 'Сметы'

    def __str__(self):
        return str(self.date_ks2)


class GIP(models.Model):
    """Главный инженер проекта"""

    fio = models.CharField(max_length = 256, verbose_name='Фамилия и инициалы', blank=True, null=True)
    organization = models.CharField(max_length = 512, verbose_name='Организация', blank=True, null=True)

    class Meta:
        verbose_name = 'ГИП'
        verbose_name_plural = 'ГИПы'

    def __str__(self):
        return str(self.fio)


class District(models.Model):
    """Районы"""

    name = models.CharField(max_length = 128, verbose_name='Район', blank=True, null=True)

    class Meta:
        verbose_name = 'Район'
        verbose_name_plural = 'Районы'

    def __str__(self):
        return str(self.name)


class Person(models.Model):
    """Работники"""

    fio = models.CharField(max_length = 256, verbose_name='ФИО', blank=True, null=True)
    post = models.CharField(max_length = 512, verbose_name='Должность', blank=True, null=True)

    class Meta:
        verbose_name = 'Должностное лицо'
        verbose_name_plural = 'Должностные лица'

    def __str__(self):
        return str(self.fio)


class Commission(models.Model):
    """Комиссия"""

    predsedatel = models.ForeignKey(Person, on_delete=models.CASCADE, verbose_name='Председатель', blank=True, null=True, related_name='predseds')
    predstavitel = models.ForeignKey(Person, on_delete=models.CASCADE, verbose_name='Представитель', blank=True, null=True, related_name='predstav')
    district = models.ForeignKey(District, on_delete=models.CASCADE, verbose_name='Район', blank=True, null=True)

    class Meta:
        verbose_name = 'Комиссия'
        verbose_name_plural = 'Комиссии'

    def __str__(self):
        return str(self.district.name)


class TehNadzor(models.Model):
    """Технадзор"""

    person = models.ForeignKey(Person, on_delete=models.CASCADE, verbose_name='Технадзор', blank=True, null=True, related_name='tehnadzor')
    district = models.ForeignKey(District, on_delete=models.CASCADE, verbose_name='Район', blank=True, null=True)

    class Meta:
        verbose_name = 'Технадзор'
        verbose_name_plural = 'Технадзоры'

    def __str__(self):
        return str(self.person)


class Order(models.Model):
    """Приказы"""

    number_date = models.CharField(max_length = 256, verbose_name='Номер и дата приказа', blank=True, null=True)
    district = models.ForeignKey(District, on_delete=models.CASCADE, verbose_name='Район', blank=True, null=True)

    class Meta:
        verbose_name = 'Приказ'
        verbose_name_plural = 'Приказы'

    def __str__(self):
        return str(self.district)


class Certificate(models.Model):
    """Сертификаты"""

    name = models.CharField(max_length = 512, verbose_name='Сертификат', blank=True, null=True)

    class Meta:
        verbose_name = 'Сертификат'
        verbose_name_plural = 'Сертификаты'

    def __str__(self):
        return str(self.name)


class Diametr(models.Model):
    """Труба, в которой есть только диаметр"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)

    class Meta:
        verbose_name = 'Диаметр'
        verbose_name_plural = 'Диаметры'

    def __str__(self):
        return str(self.diametr)


class Diametr_x(models.Model):
    """Труба, в которой есть только диаметр и Х"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    x = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_x'
        verbose_name_plural = 'Диаметры_x'


class Diametr_dlina(models.Model):
    """Труба, в которой есть только диаметр и длина"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_длина'
        verbose_name_plural = 'Диаметры_длины'


class Diametr_kolvo(models.Model):
    """Труба, в которой есть только диаметр и колво"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_колво'
        verbose_name_plural = 'Диаметры_колво'


class Diametr_dlina_kolvo(models.Model):
    """Труба, в которой есть только диаметр, длина и колво"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_длина_колво'
        verbose_name_plural = 'Диаметры_длины_колво'


class Diametr_x_dlina(models.Model):
    """Труба, в которой есть только диаметр, длина и колво"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    x = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_x_длина'
        verbose_name_plural = 'Диаметры_x_длины'


class Diametr_x_dlina_kolvo(models.Model):
    """Труба, в которой есть только диаметр, x, длина и колво"""

    diametr = models.CharField(max_length = 256, blank=True, null=True)
    x = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_x_длина_колво'
        verbose_name_plural = 'Диаметр_x_длина_колво'


class Dop_dann_sharovyi_kran(models.Model):
    """Объект для доп данных - шаровый кран"""

    kran = models.ForeignKey(Diametr, on_delete=models.CASCADE, related_name='kran', blank=True, null=True)
    mufta = models.ForeignKey(Diametr, on_delete=models.CASCADE, related_name='mufta', blank=True, null=True)

    def __str__(self):
        return str(self.kran)

    class Meta:
        verbose_name = 'Доп_данные_шаровый_кран'
        verbose_name_plural = 'Доп_данные_шаровые_краны'


class Dop_dann_gazopr_v_zashit(models.Model):
    """Объект для доп данных - В защитном футляре"""

    truba1 = models.ForeignKey(Diametr_x, on_delete=models.CASCADE, related_name='truba1', blank=True, null=True)
    truba2 = models.ForeignKey(Diametr_x, on_delete=models.CASCADE, related_name='truba2', blank=True, null=True)

    def __str__(self):
        return str(self.truba1)

    class Meta:
        verbose_name = 'Доп_данные_в_защитном_футляре'
        verbose_name_plural = 'Доп_данные_в_защитных_футлярах'


class Svarshik(models.Model):
    """Сварщики"""
    
    fio = models.CharField(max_length = 256, blank=True, null=True)
    type = models.CharField(max_length = 256, blank=True, null=True)
    diametr = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)
    date_svarki = models.DateField(verbose_name='Дата сварки', blank=True, null=True)

    def __str__(self):
        return str(self.fio)

    @property
    def date_svarki_html(self):
        if self.date_svarki:
            return datetime.datetime.strftime(self.date_svarki, "%Y-%m-%d")

    class Meta:
        verbose_name = 'Сварщик'
        verbose_name_plural = 'Сварщики'


class Cokol_soed_stal_gazoprovod(models.Model):
    """Цокольный соединитель"""

    PE = models.CharField(max_length = 256, blank=True, null=True)
    ST = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.PE)

    class Meta:
        verbose_name = 'Цокольный_соединитель'
        verbose_name_plural = 'Цокольные_соединители'


class Diametrs_3_dlina(models.Model):
    """Тройник"""

    diametr1 = models.CharField(max_length = 256, blank=True, null=True)
    diametr2 = models.CharField(max_length = 256, blank=True, null=True)
    diametr3 = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr1)

    class Meta:
        verbose_name = 'Диаметр_3_длина'
        verbose_name_plural = 'Диаметр_3_длины'


class Diametrs_3(models.Model):
    """Тройник"""

    diametr1 = models.CharField(max_length = 256, blank=True, null=True)
    diametr2 = models.CharField(max_length = 256, blank=True, null=True)
    diametr3 = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr1)

    class Meta:
        verbose_name = 'Диаметр_3'
        verbose_name_plural = 'Диаметры_3'


class Diametrs_3_kolvo(models.Model):
    """Тройник"""

    diametr1 = models.CharField(max_length = 256, blank=True, null=True)
    diametr2 = models.CharField(max_length = 256, blank=True, null=True)
    diametr3 = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr1)

    class Meta:
        verbose_name = 'Диаметр_3_колво'
        verbose_name_plural = 'Диаметры_3_колво'


class Diametr_x_dlina_prim(models.Model):
    diametr = models.CharField(max_length = 256, blank=True, null=True)
    x = models.CharField(max_length = 256, blank=True, null=True)
    dlina = models.CharField(max_length = 256, blank=True, null=True)
    prim = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.diametr)

    class Meta:
        verbose_name = 'Диаметр_x_длина_примечание'
        verbose_name_plural = 'Диаметры_x_длины_примечание'


class Neraz_soed_stal_gazoprovod(models.Model):
    PE = models.CharField(max_length = 256, blank=True, null=True)
    ST = models.CharField(max_length = 256, blank=True, null=True)
    kolvo = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.PE)

    class Meta:
        verbose_name = 'Нераз_соед_стальной_газопровод'
        verbose_name_plural = 'Нераз_соед_стальные_газопроводы'


class Podzem_polietilen_gazoprovod(models.Model):
    truba = models.ManyToManyField(Diametr_x_dlina, related_name='truba', blank=True, null=True)
    mufta =  models.ManyToManyField(Diametr_x_dlina, related_name='mufta', blank=True, null=True)
    otvod = models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='otvod', blank=True, null=True)
    troinik = models.ForeignKey(Diametrs_3_kolvo, on_delete=models.CASCADE, related_name='troinik', blank=True, null=True)
    sedelka = models.ForeignKey(Diametrs_3_kolvo, on_delete=models.CASCADE, related_name='sedelka', blank=True, null=True)
    zaglushka = models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='zaglushka', blank=True, null=True)
    kran = models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='kran', blank=True, null=True)
    lenta = models.CharField(max_length = 256, blank=True, null=True)
    znak = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.znak)

    class Meta:
        verbose_name = 'Подземный_полиэтиленовый_газопровод'
        verbose_name_plural = 'Подземные_полиэтиленовые_газопроводы'


class Nadzem_stal_gazoprovod(models.Model):
    truba = models.ManyToManyField(Diametr_x_dlina,blank=True, null=True)
    izolir_soed =  models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='izolir2', blank=True, null=True)
    kran_stal =  models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='kran2', blank=True, null=True)
    otvod =  models.ForeignKey(Diametr_kolvo, on_delete=models.CASCADE, related_name='otvod2', blank=True, null=True)
    cokolnyi_vvod = models.ForeignKey(Cokol_soed_stal_gazoprovod, on_delete=models.CASCADE, blank=True, null=True)
    kreplenie = models.CharField(max_length = 256, blank=True, null=True)
    stoika = models.ForeignKey(Diametr_dlina_kolvo, on_delete=models.CASCADE, blank=True, null=True)

    def __str__(self):
        return str(self.kreplenie)

    class Meta:
        verbose_name = 'Надземный_стальной_газопровод'
        verbose_name_plural = 'Надземные_стальные_газопроводы'


class Podzem_stal_gazoprovod(models.Model):
    truba = models.ManyToManyField(Diametr_x_dlina_prim,blank=True, null=True)
    neraz_soed =  models.ForeignKey(Neraz_soed_stal_gazoprovod, on_delete=models.CASCADE, blank=True, null=True)
    kontrolnaya_trubka = models.CharField(max_length = 256, blank=True, null=True)
    otvod_90 = models.CharField(max_length = 256, blank=True, null=True)
    opoznavat_znak = models.CharField(max_length = 256, blank=True, null=True)

    def __str__(self):
        return str(self.otvod_90)

    class Meta:
        verbose_name = 'Подземный_стальной_газопровод'
        verbose_name_plural = 'Подземные_стальные_газопроводы'


class Kontragent(models.Model):
    name_kontragent = models.CharField(max_length = 512, verbose_name='Имя контрагента', blank=True, null=True)
    INN = models.CharField(max_length = 256, verbose_name='ИНН контрагента', blank=True, null=True)
    KPP = models.CharField(max_length = 256, verbose_name='КПП контрагента', blank=True, null=True)
    Ur_address = models.CharField(max_length = 1024, verbose_name='Юридический адрес', blank=True, null=True)
    telephone = models.CharField(max_length = 128, verbose_name='Телефон', blank=True, null=True)
    podpisant = models.ForeignKey(Person, on_delete=models.CASCADE, verbose_name='Подписант ген подрядчика', blank=True, null=True)

    def __str__(self):
        return str(self.name_kontragent)

    class Meta:
        verbose_name = 'Контрагент'
        verbose_name_plural = 'Контрагенты'


class Rostehnadzor(models.Model):
    pk1 = models.CharField(max_length = 128, verbose_name='Номер 1 ПК', blank=True, null=True)
    pk1_diam = models.CharField(max_length = 128, verbose_name='Диаметр 1 ПК', blank=True, null=True)
    pk2 = models.CharField(max_length = 128, verbose_name='Номер 2 ПК', blank=True, null=True)
    pk2_diam = models.CharField(max_length = 128, verbose_name='Диаметр 2 ПК', blank=True, null=True)

    truba =  models.ForeignKey(Diametr_x, on_delete=models.CASCADE, verbose_name='Труба', blank=True, null=True)

    def __str__(self):
        if self.pk1 and self.pk1_diam and self.pk2 and self.pk2_diam:
            return str(f'ПК{self.pk1}+{self.pk1_diam}-ПК{self.pk2}+{self.pk2_diam}')
        if self.pk1:
            return self.pk1
        if self.pk2:
            return self.pk2
        if self.pk1_diam:
            return self.pk1_diam
        if self.pk2_diam:
            return self.pk2_diam
        return ''

    class Meta:
        verbose_name = 'Ростехнадзор'
        verbose_name_plural = 'Ростехнадзоры'


class Prodavlivanie(models.Model):
    pk0_1_diam = models.CharField(max_length = 128, verbose_name='ПК 0 1 диаметр', blank=True, null=True)
    pk0_2_diam = models.CharField(max_length = 128, verbose_name='ПК 0 2 диаметр', blank=True, null=True)

    truba =  models.ForeignKey(Diametr_x_dlina, on_delete=models.CASCADE, verbose_name='Труба', blank=True, null=True)

    def __str__(self):
        if self.pk0_1_diam and self.pk0_2_diam:
            return str(f'ПК0+{self.pk0_1_diam}-ПК0+{self.pk0_2_diam}')
        if self.pk0_1_diam:
            return str(f'ПК0+{self.pk0_1_diam}-ПК0+_')
        if self.pk0_2_diam:
            return str(f'ПК0+_-ПК0+{self.pk0_2_diam}')
        return ''

    class Meta:
        verbose_name = 'Продавливание'
        verbose_name_plural = 'Продавливание'


class Object(models.Model):
    name_object = models.CharField(max_length = 1024, verbose_name='Имя объекта', blank=True, null=True)
    zakazchik = models.CharField(max_length = 1024, verbose_name='Заказчик', blank=True, null=True)
    place = models.CharField(max_length = 1024, verbose_name='Местоположение', blank=True, null=True)
    kontragent = models.ForeignKey(Kontragent, on_delete=models.CASCADE, verbose_name='Контрагент', blank=True, null=True)
    Nomer_razm = models.CharField(max_length = 256, verbose_name='Номер договора на размещение', blank=True, null=True)
    Data_razm = models.DateField(verbose_name='Дата договора на размещение', blank=True, null=True)
    Nomer_proekt = models.CharField(max_length = 256, verbose_name='Номер проекта', blank=True, null=True)
    Nomer_zadaniya = models.CharField(max_length = 256, verbose_name='Номер задания', blank=True, null=True)
    Date_proekt = models.DateField(verbose_name='Дата проектирования', blank=True, null=True)
    gip = models.ForeignKey(GIP, on_delete=models.CASCADE, verbose_name='ГИП', blank=True, null=True)
    district = models.ForeignKey(District, on_delete=models.CASCADE, verbose_name='Район', blank=True, null=True)
    kod_object = models.CharField(max_length = 1024, verbose_name='Код объекта', blank=True, null=True)
    zayv = models.CharField(max_length = 1024, verbose_name='Заявитель', blank=True, null=True)
    ks2_zakazchik = models.ForeignKey(Person, on_delete=models.CASCADE,
        verbose_name='Заказчик', blank=True, null=True, related_name='ks2_zakazchik')
    ks2_podryadchik= models.ForeignKey(Person, on_delete=models.CASCADE,
        verbose_name='Подрядчик', blank=True, null=True, related_name='ks2_podryadchik')
    ks2_object = models.CharField(max_length = 1024, verbose_name='Объект', blank=True, null=True)
    ks2_stroika = models.CharField(max_length = 1024, verbose_name='Стройка', blank=True, null=True)
    ks11_predsedatel = models.ForeignKey(Person, on_delete=models.CASCADE,
        verbose_name='Председатель', blank=True, null=True, related_name='ks11_predsedatel')
    ks11_predstav_ekspl = models.ForeignKey(Person, on_delete=models.CASCADE,  
        verbose_name='Представитель экспл.', blank=True, null=True, related_name='ks11_predstav_ekspl')
    ks11_predstav_proekt = models.ForeignKey(Person, on_delete=models.CASCADE,
        verbose_name='Представитель проект.', blank=True, null=True, related_name='ks11_predstav_proekt')
    prover_davl = models.CharField(max_length = 256, verbose_name='Проверочное давление', blank=True, null=True)
    davlenie_name = models.CharField(max_length = 256, verbose_name='Название давления', blank=True, null=True)
    davl = models.CharField(max_length = 256, verbose_name='Давление', blank=True, null=True)
    gazoprovod_podzem_stal =  models.ForeignKey(Podzem_stal_gazoprovod, on_delete=models.CASCADE,
        verbose_name='Подземный стальной газопровод', blank=True, null=True, related_name='gazoprovod_podzem_stal')
    gazoprovod_podzem_poliet =  models.ForeignKey(Podzem_polietilen_gazoprovod, on_delete=models.CASCADE,
        verbose_name='Подземный полиэтиленовый газопровод', blank=True, null=True, related_name='gazoprovod_podzem_poliet')
    gazoprovod_nadzem_stal =  models.ForeignKey(Nadzem_stal_gazoprovod, on_delete=models.CASCADE, 
        verbose_name='Надземный стальной газопровод', blank=True, null=True, related_name='gazoprovod_nadzem_stal')
    certificates =  models.ManyToManyField(Certificate, verbose_name='Сертификаты', blank=True, null=True, related_name='cert')
    gazoprovod = models.BooleanField(default=False)
    prodavlivanie = models.BooleanField(default=False)
    svarshik1 = models.ForeignKey(Svarshik, on_delete=models.CASCADE, verbose_name='Сварщик 1', blank=True, null=True, related_name='svar1')
    svarshik2 = models.ForeignKey(Svarshik, on_delete=models.CASCADE, verbose_name='Сварщик 2', blank=True, null=True, related_name='svar2')
    tehnadzor = models.ForeignKey(TehNadzor, on_delete=models.CASCADE, verbose_name='Технадзор', blank=True, null=True)
    proektnaya_org = models.CharField(max_length = 256, verbose_name='Проектная организация', blank=True, null=True, default='')
    Data_zamera1 = models.DateField(verbose_name='Дата 1 замера', blank=True, null=True)
    Data_zamera2 = models.DateField(verbose_name='Дата 2 замера', blank=True, null=True)
    Data_sost_project = models.DateField(verbose_name='Дата составления проекта', blank=True, null=True)
    Data_razbiv = models.DateField(verbose_name='Дата разбивки', blank=True, null=True)
    Data_produv = models.DateField(verbose_name='Дата продувки', blank=True, null=True)
    Data_ukl = models.DateField(verbose_name='Дата укладки', blank=True, null=True)
    zashitnyi_futlyar = models.ForeignKey(Diametr_x_dlina, on_delete=models.CASCADE, verbose_name='Защитный футляр', blank=True, null=True)
    futlyar_na_vyhode = models.ForeignKey(Diametr_x_dlina_kolvo, on_delete=models.CASCADE, verbose_name='Футляр на выходе земли', blank=True, null=True)
    opora = models.ForeignKey(Diametr_x_dlina_kolvo, on_delete=models.CASCADE, 
        verbose_name='Опора под газопровод', blank=True, null=True, related_name='favorites')
    dop_dann_sharovyi_kran = models.ManyToManyField(Dop_dann_sharovyi_kran, verbose_name='Доп', blank=True, null=True, related_name='shar')
    dop_dann_gazopr_v_zashit = models.ManyToManyField(Dop_dann_gazopr_v_zashit, verbose_name='Доп', blank=True, null=True, related_name='gaz')
    dop_dann_futlyar_na_vyhode = models.ManyToManyField(Diametr_x, verbose_name='Доп', blank=True, null=True, related_name='na_vyh')
    dop_dann_ob_ustanovke_opor = models.ManyToManyField(Diametr_x, verbose_name='Доп', blank=True, null=True, related_name='opor')
    smeta = models.ForeignKey(Smeta, on_delete=models.CASCADE, verbose_name='Данные для сметы', blank=True, null=True)
    rostehnadzor = models.ForeignKey(Rostehnadzor, on_delete=models.CASCADE, verbose_name='Ростехнадзор', blank=True, null=True)
    prodavl = models.ForeignKey(Prodavlivanie, on_delete=models.CASCADE, verbose_name='Продавливание', blank=True, null=True)

    class Meta:
        verbose_name = 'Объект'
        verbose_name_plural = 'Объекты'

    def __str__(self):
        return str(self.name_object)

    @property
    def full_kontragent(self):
        if self.kontragent:
            if self.kontragent.name_kontragent or self.kontragent.INN or self.kontragent.KPP or self.kontragent.Ur_address:
                return f"{self.kontragent.name_kontragent}, ИНН {self.kontragent.INN}, КПП {self.kontragent.KPP}; {self.kontragent.Ur_address};"
            else:
                return ''

    @property
    def order(self):
        try:
            return Order.objects.get(district = self.district.pk).number_date
        except:
            return ''


    @property
    def date_nachala_rabot(self):
        if self.smeta:
            if self.smeta.date_nach_rabot:
                return datetime.date.strftime(self.smeta.date_nach_rabot,"%m.%Y")
            else:
                return ''
        else:
            return ''

    @property
    def date_ks2_2(self):
        if self.smeta:
            if self.smeta.date_ks2:
                return datetime.date.strftime(self.smeta.date_ks2,"%m.%Y")
            else:
                return ''
        else:
            return ''

    @property
    def name_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Газопровод низкого давления P = {} МПа.".format(self.davl)
        elif self.prover_davl == 0.6:
            return "Газопровод среднего давления P = {} МПа.".format(self.davl)
        elif self.davlenie_name or self.davl:
            return "Газопровод {} P = {} МПа.".format(self.davlenie_name,self.davl)
        else:
            return 'Газопровод'

    @property
    def name_podzem_stal_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Подземный стальной газопровод низкого давления:".format(self.davl)
        elif self.prover_davl == 0.6:
            return "Подземный стальной газопровод среднего давления:".format(self.davl)
        elif self.davlenie_name or self.davl:
            return "Подземный стальной газопровод {}:".format(self.davlenie_name,self.davl)
        else:
            return 'Подземный стальной газопровод'

    @property
    def full_name_podzem_stal_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Подземный стальной газопровод низкого давления       "
        elif self.prover_davl == 0.6:
            return "Подземный стальной газопровод среднего давления        "
        elif self.davlenie_name:
            return "Подземный стальной газопровод {}        ".format(self.davlenie_name)
        else:
            return 'Подземный стальной газопровод'

    @property
    def name_podzem_poliet_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Подземный полиэтиленовый газопровод низкого давления:".format(self.davl)
        elif self.prover_davl == 0.6:
            return "Подземный полиэтиленовый газопровод среднего давления:".format(self.davl)
        elif self.davlenie_name or self.davl:
            return "Подземный полиэтиленовый газопровод {}:".format(self.davlenie_name,self.davl)
        else:
            return 'Подземный полиэтиленовый газопровод'

    @property
    def full_name_podzem_poliet_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Подземный полиэтиленовый газопровод низкого давления ПЭ100 SDR11     "
        elif self.prover_davl == 0.6:
            return "Подземный полиэтиленовый газопровод низкого давления ПЭ100 SDR11     "
        elif self.davlenie_name:
            return "Подземный полиэтиленовый газопровод {} ПЭ100 SDR11     ".format(self.davlenie_name)
        else:
            return 'Подземный полиэтиленовый газопровод'


    @property
    def name_nadzem_stal_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Надземный стальной газопровод низкого давления:".format(self.davl)
        elif self.prover_davl == 0.6:
            return "Надземный стальной газопровод среднего давления:".format(self.davl)
        elif self.davlenie_name or self.davl:
            return "Надземный стальной газопровод {}:".format(self.davlenie_name,self.davl)
        else:
            return "Надземный стальной газопровод"

    @property
    def full_name_nadzem_stal_gazoprovod(self):
        if self.prover_davl == 0.3:
            return "Надземный стальной газопровод низкого давления       "
        elif self.prover_davl == 0.6:
            return "Надземный стальной газопровод среднего давления        "
        elif self.davlenie_name:
            return "Надземный стальной газопровод {}        ".format(self.davlenie_name)
        else:
            return "Надземный стальной газопровод"

    @property
    def dlina_podzem_stal(self):
        itogo = 0.0
        if self.gazoprovod_podzem_stal:
            for truba in self.gazoprovod_podzem_stal.truba.all():
                if truba.dlina:
                    dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)
        return itogo

    @property
    def dlina_podzem_poliet(self):
        itogo = 0.0
        if self.gazoprovod_podzem_poliet:
            for truba in self.gazoprovod_podzem_poliet.truba.all():
                if truba.dlina:
                    dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)
        return itogo

    @property
    def dlina_nadzem_stal(self):
        itogo = 0.0
        if self.gazoprovod_nadzem_stal:
            for truba in self.gazoprovod_nadzem_stal.truba.all():
                if truba.dlina:
                    dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)
        return itogo

    @property
    def itogo(self):
        itogo = 0.0
        if self.gazoprovod_podzem_stal:

            for truba in self.gazoprovod_podzem_stal.truba.all():
                if truba.dlina:
                    dlina = truba.dlina
                    if ',' in truba.dlina:
                        dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)

        if self.gazoprovod_podzem_poliet:

            for truba in self.gazoprovod_podzem_poliet.truba.all():
                if truba.dlina:
                    dlina = truba.dlina
                    if ',' in truba.dlina:
                        dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)

        if self.gazoprovod_nadzem_stal:
            for truba in self.gazoprovod_nadzem_stal.truba.all():
                if truba.dlina:
                    dlina = truba.dlina
                    if ',' in truba.dlina:
                        dlina = truba.dlina.replace(',','.')
                    itogo += float(dlina)

        return float(itogo)
