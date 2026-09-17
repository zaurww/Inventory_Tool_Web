ANBAR UÇOTU — VERSİYA 1.2.2

BAŞLAMAQ

Veb versiya: https://zaurww.github.io/Inventory_Tool_Web/
Saytdan istifadə üçün Python quraşdırmaq lazım deyil.
Versiya səhifənin aşağı hissəsində və hesabatın Məlumat vərəqində göstərilir.
Kompüterdə yerli sınaq üçün aşağıdakı Run_Web_Pilot.bat üsulu da qalır.

İlk istifadə üçün “Şablonu endir” düyməsi ilə boş giriş kitabını endirin.
Excel-də Parametrlər vərəqində şirkəti və hər iki ayı, sonra malları,
kontragentləri və əməliyyatları doldurun. Kitabı saxlayıb saytda seçin.
“Demonu endir” düyməsi ilə uydurma məlumatlı sınaq kitabını endirə bilərsiniz.
Demonu saytda hesablayın və hesabatı “Nümunə yoxlaması” vərəqindəki nəticələrlə
müqayisə edin. Hesabat məbləğləri 2 onluq rəqəmə yuvarlaq göstərir;
yoxlama vərəqində dəqiq qiymətlər verilib. Demo yalnız sınaq üçündür.
Hər iki fayl hesablama modulunun yüklənməsini gözləmədən endirilə bilər.

Run_Web_Pilot.bat faylını açın. Brauzerdə http://127.0.0.1:8765/ açılır.
XLSX kitabını seçin və “Hesabla və faylları hazırla” düyməsini basın.
Hesabatı və *_AZ.xlsx adlı yenilənmiş giriş kitabını endirin.
İşi məhz yenilənmiş giriş kitabında davam etdirin: daimi sətir kodları ondadır.
İlkin seçilmiş fayl dəyişdirilmir.

Uçot məlumatları brauzerin yaddaşında hesablanır, serverə göndərilmir.
Proqram modullarının ilk yüklənməsi üçün internet lazımdır.
Brauzeri bağladıqda endirilməmiş nəticələr itir. İş kitabının və hesabatların
ehtiyat nüsxələrini ayrıca saxlayın. Hər bazanı bir məsul şəxs aparmalıdır.

FORMAT VƏ UYĞUNLUQ

Rus və ingilis dilində əvvəlki kitablar oxunur və ayrıca AZ nüsxəsinə çevrilir.
Malların, kontragentlərin və sənədlərin adları, məbləğlər və kodlar tərcümə edilmir.
Yeni kitabda ağıllı cədvəllər yoxdur. Adi diapazonlar, filtrlər və seçim siyahıları var.
Adi məlumat sətrlərinin hündürlüyü 18 punktdur. Uzun mətnli sətrlər məzmuna görə böyüyür.
Əvvəl endirilmiş kitabı saytda hesablayıb yenilənmiş nüsxəsini endirin.
Gizli _InventoryMeta vərəqi format versiyasını saxlayır. Onu dəyişməyin.
Naməlum format versiyası olduqda proqram hesablamanı dayandırır.
Köhnə kitabdakı fərdi formulalar və adlandırılmış diapazonlar avtomatik
çevirmədən əvvəl ayrıca yoxlanmalıdır.

MƏLUMATLARIN DAXİL EDİLMƏSİ

Başlıqlar 6-cı, məlumatlar 7-ci sətirdən başlayır. Başlıqları dəyişməyin.
Ulduz (*) vacib sahəni göstərir. Giriş sahələrində formula deyil, dəyər olmalıdır.
Yalnız dəyərləri yapışdırın ki, başqa faylın formatı köçürülməsin.
Bütün diapazonu birlikdə sıralayın; bir sütunu ayrıca sıralamayın.
Seçim siyahıları və filtrlər ən azı 10 000 giriş sətrini əhatə edir.
Bu həddi keçdikdə yenidən hesablayın və yenilənmiş kitabı endirin.
Siyahılarda boş sətrlər görünə bilər. Python bütün dolu sətrləri yoxlayır.

Parametrlər:
  Şirkəti, uçotun başlanğıc ayını və hesabatın son ayını göstərin.
  Hər iki tarix ayın ilk günü olmalıdır. Başlanğıc qalıq sıfırdır.

Mallar:
  Hər SKU bir dəfə yazılır. Malın adı və ölçü vahidi vacibdir.
  Eyni SKU müxtəlif alış və satış sətrlərində təkrarlana bilər.
  Böyük/kiçik hərflər və kənar boşluqlar kodları fərqləndirmir.

Kontragentlər:
  Təchizatçı, alıcı və xidmət göstərənləri bir dəfə qeyd edin.
  Eyni kontragent bir neçə rolda iştirak edə bilər.

Alışlar:
  Alış növü İdxal və ya Yerli olur. Miqdar müsbət olmalıdır.
  İdxalda bəyannamə, yerli alışda e-qaimə nömrəsini göstərin.
  Bir tədarükün bütün mal sətrlərində eyni nömrə yazılır.
  Müstəqil tədarüklər üçün eyni nömrəni istifadə etməyin.
  Məbləğ vahidin qiyməti deyil, bütün sətrin ƏDV-siz məbləğidir.
  AZN məzənnəsi 1 olmalıdır. Yeni əməliyyatın sətir kodunu boş saxlayın.

Xərclər:
  Hər xərc ayrı sətirdə yazılır və tədarük nömrəsinə bağlanır.
  Xərclər malların AZN ilə alış dəyərinə mütənasib bölüşdürülür.
  Bütün məlum xərclər, sənəd tarixindən asılı olmayaraq daxil edilir.
  Gec xərclər ilkin alışın və sonrakı ayların nəticələrini yenidən hesablayır.
  Mənfi məbləğ əvvəlki xərcin düzəlişidir.

Satışlar:
  Əməliyyat Satış və ya Silinmə olur. Miqdar həmişə müsbətdir.
  Satış üçün alıcı və ƏDV-siz satış məbləği vacibdir; sıfır məbləği 0 yazın.
  Silinmədə satış məbləği boş və ya 0 olmalıdır.

Qaytarmalar:
  Əvvəlcə ilkin əməliyyatı hesablayaraq onun daimi kodunu alın.
  Alıcı qaytarması satışa, təchizatçıya qaytarma alışa bağlanır.
  Qaytarma ilkin tarixdən əvvəl ola və ilkin miqdarı aşa bilməz.
  Alıcı qaytarmasında satış məbləğinin azalması, o cümlədən 0, göstərilməlidir.
  Təchizatçıya qaytarmada bu sahə 0 və ya boş olmalıdır.

KODLAR VƏ YOXLAMALAR

P — alış, S — satış/silinmə, E — xərc, R — qaytarma kodudur.
Mövcud kodu dəyişməyin. Sətri yeni əməliyyat üçün kopyaladıqda kodu təmizləyin.
Kod təkrarı xətadır. Oxşar əməliyyatlar xəbərdarlıq yaradır, lakin hər iki
sətir hesablamaya daxil edilir. Təkrarları proqram özü silmir.
Xəta olduqda hesabat yaradılmır. Yoxlamalar faylını endirin və səbəbləri düzəldin.
Çevirmə tamamlanıbsa, kodları olan AZ giriş kitabı xəta zamanı da endirilə bilər.

HESABLAMA VƏ HESABATLAR

Aylıq orta maya dəyəri (Monthly AVCO) istifadə edilir.
Ay ərzindəki bütün satışlar eyni orta maya dəyəri ilə qiymətləndirilir.
Hərəkətli orta maya dəyəri (Moving Average) hələ dəstəklənmir.
Alıcı qaytarması ilkin satış ayının maya dəyəri ilə bərpa edilir.
Təchizatçıya qaytarma ilkin alışın xərclər daxil tam maya dəyəri ilə çıxılır.
Bu məbləğ təchizatçının pul geriödənişi deyil, anbar dəyəridir.

Aylıq yekun — satış məbləği, maya dəyəri, mənfəət, silinmə və qalıq.
Aylıq anbar uçotu — hər mal üzrə aylıq miqdar və dəyər hərəkətləri.
Satışların təfərrüatı və Satışların yekunu — satışlar və alıcı qaytarmaları.
Tədarükün maya dəyəri — alış dəyəri və bölüşdürülmüş xərclər.
Tədarüklərin yekunu — bəyannamə / e-qaimə üzrə yekunlar.
Tədarük xərcləri — xərclərin sənədlər və icraçılar üzrə təfərrüatı.
Silinmələr — ayrıca silinmə reyestri.
Yoxlamalar — xətalar, xəbərdarlıqlar və dövr məlumatları.
Məlumat — proqram versiyası, mənbə faylı, vaxt və hesablama qaydaları.

Hesabatlar formulasız, statik Excel dəyərləridir. Məlumat dəyişdikdə yenidən hesablayın.
Məbləğlər 2, maya dəyəri 6 onluq rəqəmlə göstərilir.
Aralıq hesablamalar yuvarlaqlaşdırılmır. Mənfi qalıq ayın sonunda yoxlanılır.
Ümumi mənfəət xalis mənfəət deyil: inzibati və digər əməliyyat xərcləri daxil deyil.
Bir ümumi anbar, AZN və ƏDV-siz məbləğlər istifadə edilir.
