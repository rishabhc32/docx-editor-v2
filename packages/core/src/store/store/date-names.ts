// Month and weekday names for a date control's language (`w:lid`), as Word for Mac shows them in
// its date pictures (2026-11-25). Fixed data, never the runtime's locale data: two copies on
// different runtimes must write the same text. A language Word shows in English (th) has no row.

/** One language's names, as Word shows a date picture in that language. */
export interface DateNames {
  /** Month names, January first (`MMMM`). */
  readonly months: readonly string[];
  /** Short month names (`MMM`). */
  readonly shortMonths: readonly string[];
  /** Weekday names, Sunday first (`dddd`). */
  readonly weekdays: readonly string[];
  /** Short weekday names (`ddd`). */
  readonly shortWeekdays: readonly string[];
  /** Month names after a day number (`d MMMM`): the genitive where the language has one. */
  readonly monthsAfterDay: readonly string[];
  /** Whether a day number after the month (`MMMM d`) also takes {@link monthsAfterDay}. */
  readonly dayEitherSide: boolean;
  /** The text of day 1 right before a month name (`d MMMM`), where it is no plain number. */
  readonly firstDay: string | undefined;
}

/**
 * The rows: a language tag in lower case, then `|`-joined months, short months, weekdays and
 * short weekdays, and the months after a day number where they differ. A region has a row only
 * where Word's names for it differ from its language's.
 */
const ROWS: readonly (readonly string[])[] = [
  [
    'en',
    'January|February|March|April|May|June|July|August|September|October|November|December',
    'Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec',
    'Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday',
    'Sun|Mon|Tue|Wed|Thu|Fri|Sat',
  ],
  [
    'de',
    'Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember',
    'Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez',
    'Sonntag|Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag',
    'So|Mo|Di|Mi|Do|Fr|Sa',
  ],
  [
    'de-at',
    'Jänner|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember',
    'Jän|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez',
    'Sonntag|Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag',
    'So|Mo|Di|Mi|Do|Fr|Sa',
  ],
  [
    'fr',
    'janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre',
    'janv.|févr.|mars|avr.|mai|juin|juil.|août|sept.|oct.|nov.|déc.',
    'dimanche|lundi|mardi|mercredi|jeudi|vendredi|samedi',
    'dim.|lun.|mar.|mer.|jeu.|ven.|sam.',
  ],
  [
    'fr-ca',
    'janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre',
    'janv.|févr.|mars|avr.|mai|juin|juill.|août|sept.|oct.|nov.|déc.',
    'dimanche|lundi|mardi|mercredi|jeudi|vendredi|samedi',
    'dim.|lun.|mar.|mer.|jeu.|ven.|sam.',
  ],
  [
    'es',
    'enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre',
    'ene|feb|mar|abr|may|jun|jul|ago|sept|oct|nov|dic',
    'domingo|lunes|martes|miércoles|jueves|viernes|sábado',
    'dom|lun|mar|mié|jue|vie|sáb',
  ],
  [
    'es-mx',
    'enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre',
    'ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic',
    'domingo|lunes|martes|miércoles|jueves|viernes|sábado',
    'dom|lun|mar|mié|jue|vie|sáb',
  ],
  [
    'it',
    'gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre',
    'gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic',
    'domenica|lunedì|martedì|mercoledì|giovedì|venerdì|sabato',
    'dom|lun|mar|mer|gio|ven|sab',
  ],
  [
    'nl',
    'januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december',
    'jan|feb|mrt|apr|mei|jun|jul|aug|sep|okt|nov|dec',
    'zondag|maandag|dinsdag|woensdag|donderdag|vrijdag|zaterdag',
    'zo|ma|di|wo|do|vr|za',
  ],
  [
    'pt',
    'janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro',
    'jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez',
    'domingo|segunda-feira|terça-feira|quarta-feira|quinta-feira|sexta-feira|sábado',
    'dom|seg|ter|qua|qui|sex|sáb',
  ],
  [
    'sv',
    'januari|februari|mars|april|maj|juni|juli|augusti|september|oktober|november|december',
    'jan.|feb.|mars|apr.|maj|juni|juli|aug.|sep.|okt.|nov.|dec.',
    'söndag|måndag|tisdag|onsdag|torsdag|fredag|lördag',
    'sön|mån|tis|ons|tors|fre|lör',
  ],
  [
    'da',
    'januar|februar|marts|april|maj|juni|juli|august|september|oktober|november|december',
    'jan|feb|mar|apr|maj|jun|jul|aug|sep|okt|nov|dec',
    'søndag|mandag|tirsdag|onsdag|torsdag|fredag|lørdag',
    'søn.|man.|tirs.|ons.|tors.|fre.|lør.',
  ],
  [
    'nb',
    'januar|februar|mars|april|mai|juni|juli|august|september|oktober|november|desember',
    'jan|feb|mar|apr|mai|jun|jul|aug|sep|okt|nov|des',
    'søndag|mandag|tirsdag|onsdag|torsdag|fredag|lørdag',
    'søn.|man.|tir.|ons.|tor.|fre.|lør.',
  ],
  [
    'nn',
    'januar|februar|mars|april|mai|juni|juli|august|september|oktober|november|desember',
    'jan|feb|mar|apr|mai|jun|jul|aug|sep|okt|nov|des',
    'søndag|måndag|tysdag|onsdag|torsdag|fredag|laurdag',
    'søn|mån|tys|ons|tor|fre|lau',
  ],
  [
    'fi',
    'tammikuu|helmikuu|maaliskuu|huhtikuu|toukokuu|kesäkuu|heinäkuu|elokuu|syyskuu|lokakuu|marraskuu|joulukuu',
    'tammi|helmi|maalis|huhti|touko|kesä|heinä|elo|syys|loka|marras|joulu',
    'sunnuntai|maanantai|tiistai|keskiviikko|torstai|perjantai|lauantai',
    'su|ma|ti|ke|to|pe|la',
  ],
  [
    'is',
    'janúar|febrúar|mars|apríl|maí|júní|júlí|ágúst|september|október|nóvember|desember',
    'jan.|feb.|mar.|apr.|maí|jún.|júl.|ágú.|sep.|okt.|nóv.|des.',
    'sunnudagur|mánudagur|þriðjudagur|miðvikudagur|fimmtudagur|föstudagur|laugardagur',
    'sun.|mán.|þri.|mið.|fim.|fös.|lau.',
  ],
  [
    'pl',
    'styczeń|luty|marzec|kwiecień|maj|czerwiec|lipiec|sierpień|wrzesień|październik|listopad|grudzień',
    'sty|lut|mar|kwi|maj|cze|lip|sie|wrz|paź|lis|gru',
    'niedziela|poniedziałek|wtorek|środa|czwartek|piątek|sobota',
    'niedz.|pon.|wt.|śr.|czw.|pt.|sob.',
    'stycznia|lutego|marca|kwietnia|maja|czerwca|lipca|sierpnia|września|października|listopada|grudnia',
  ],
  [
    'cs',
    'leden|únor|březen|duben|květen|červen|červenec|srpen|září|říjen|listopad|prosinec',
    'led|úno|bře|dub|kvě|čvn|čvc|srp|zář|říj|lis|pro',
    'neděle|pondělí|úterý|středa|čtvrtek|pátek|sobota',
    'ne|po|út|st|čt|pá|so',
    'ledna|února|března|dubna|května|června|července|srpna|září|října|listopadu|prosince',
  ],
  [
    'sk',
    'január|február|marec|apríl|máj|jún|júl|august|september|október|november|december',
    'jan|feb|mar|apr|máj|jún|júl|aug|sep|okt|nov|dec',
    'nedeľa|pondelok|utorok|streda|štvrtok|piatok|sobota',
    'ne|po|ut|st|št|pi|so',
    'januára|februára|marca|apríla|mája|júna|júla|augusta|septembra|októbra|novembra|decembra',
  ],
  [
    'sl',
    'januar|februar|marec|april|maj|junij|julij|avgust|september|oktober|november|december',
    'jan.|feb.|mar.|apr.|maj|jun.|jul.|avg.|sep.|okt.|nov.|dec.',
    'nedelja|ponedeljek|torek|sreda|četrtek|petek|sobota',
    'ned.|pon.|tor.|sre.|čet.|pet.|sob.',
  ],
  [
    'hr',
    'siječnja|veljače|ožujka|travnja|svibnja|lipnja|srpnja|kolovoza|rujna|listopada|studenog|prosinca',
    'sij.|vlj.|ožu.|tra.|svi.|lip.|srp.|kol.|ruj.|lis.|stu.|pro.',
    'nedjelja|ponedjeljak|utorak|srijeda|četvrtak|petak|subota',
    'ned.|pon.|uto.|sri.|čet.|pet.|sub.',
  ],
  [
    'sr',
    'januar|februar|mart|april|maj|jun|jul|avgust|septembar|oktobar|novembar|decembar',
    'jan|feb|mar|apr|maj|jun|jul|avg|sep|okt|nov|dec',
    'nedelja|ponedeljak|utorak|sreda|četvrtak|petak|subota',
    'ned|pon|uto|sre|čet|pet|sub',
  ],
  [
    'sr-cyrl',
    'јануар|фебруар|март|април|мај|јун|јул|август|септембар|октобар|новембар|децембар',
    'јан|феб|мар|апр|мај|јун|јул|авг|сеп|окт|нов|дец',
    'недеља|понедељак|уторак|среда|четвртак|петак|субота',
    'нед|пон|уто|сре|чет|пет|суб',
  ],
  [
    'bs',
    'januar|februar|mart|april|maj|juni|juli|august|septembar|oktobar|novembar|decembar',
    'jan|feb|mar|apr|maj|jun|jul|aug|sep|okt|nov|dec',
    'nedjelja|ponedjeljak|utorak|srijeda|četvrtak|petak|subota',
    'ned|pon|uto|sri|čet|pet|sub',
  ],
  [
    'hu',
    'január|február|március|április|május|június|július|augusztus|szeptember|október|november|december',
    'jan.|febr.|márc.|ápr.|máj.|jún.|júl.|aug.|szept.|okt.|nov.|dec.',
    'vasárnap|hétfő|kedd|szerda|csütörtök|péntek|szombat',
    'V|H|K|Sze|Cs|P|Szo',
  ],
  [
    'ro',
    'ianuarie|februarie|martie|aprilie|mai|iunie|iulie|august|septembrie|octombrie|noiembrie|decembrie',
    'ian.|feb.|mar.|apr.|mai|iun.|iul.|aug.|sept.|oct.|nov.|dec.',
    'duminică|luni|marți|miercuri|joi|vineri|sâmbătă',
    'dum.|lun.|mar.|mie.|joi.|vin.|sâm.',
  ],
  [
    'bg',
    'януари|февруари|март|април|май|юни|юли|август|септември|октомври|ноември|декември',
    'яну|фев|март|апр|май|юни|юли|авг|сеп|окт|ное|дек',
    'неделя|понеделник|вторник|сряда|четвъртък|петък|събота',
    'нд|пн|вт|ср|чт|пт|сб',
  ],
  [
    'ru',
    'январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь',
    'янв.|февр.|март|апр.|май|июнь|июль|авг.|сент.|окт.|нояб.|дек.',
    'воскресенье|понедельник|вторник|среда|четверг|пятница|суббота',
    'вс|пн|вт|ср|чт|пт|сб',
    'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря',
  ],
  [
    'uk',
    'січень|лютий|березень|квітень|травень|червень|липень|серпень|вересень|жовтень|листопад|грудень',
    'січ.|лют.|бер.|квіт.|трав.|черв.|лип.|серп.|вер.|жовт.|лист.|груд.',
    'неділя|понеділок|вівторок|середа|четвер|пʼятниця|субота',
    'нд|пн|вт|ср|чт|пт|сб',
    'січня|лютого|березня|квітня|травня|червня|липня|серпня|вересня|жовтня|листопада|грудня',
  ],
  [
    'be',
    'студзень|люты|сакавік|красавік|май|чэрвень|ліпень|жнівень|верасень|кастрычнік|лістапад|снежань',
    'сту|лют|сак|кра|май|чэр|ліп|жні|вер|кас|ліс|сне',
    'нядзеля|панядзелак|аўторак|серада|чацвер|пятніца|субота',
    'нд|пн|аў|ср|чц|пт|сб',
    'студзеня|лютыа|сакавіка|красавіка|мая|чэрвеня|ліпеня|жнівеня|верасеня|кастрычніка|лістапада|снежаня',
  ],
  [
    'el',
    'Ιανουάριος|Φεβρουάριος|Μάρτιος|Απρίλιος|Μάιος|Ιούνιος|Ιούλιος|Αύγουστος|Σεπτέμβριος|Οκτώβριος|Νοέμβριος|Δεκέμβριος',
    'Ιαν|Φεβ|Μάρ|Απρ|Μάι|Ιούν|Ιούλ|Αύγ|Σεπ|Οκτ|Νοέ|Δεκ',
    'Κυριακή|Δευτέρα|Τρίτη|Τετάρτη|Πέμπτη|Παρασκευή|Σάββατο',
    'Κυρ|Δευ|Τρί|Τετ|Πέμ|Παρ|Σάβ',
    'Ιανουαρίου|Φεβρουαρίου|Μαρτίου|Απριλίου|Μαΐου|Ιουνίου|Ιουλίου|Αυγούστου|Σεπτεμβρίου|Οκτωβρίου|Νοεμβρίου|Δεκεμβρίου',
  ],
  [
    'tr',
    'Ocak|Şubat|Mart|Nisan|Mayıs|Haziran|Temmuz|Ağustos|Eylül|Ekim|Kasım|Aralık',
    'Oca|Şub|Mar|Nis|May|Haz|Tem|Ağu|Eyl|Eki|Kas|Ara',
    'Pazar|Pazartesi|Salı|Çarşamba|Perşembe|Cuma|Cumartesi',
    'Paz|Pzt|Sal|Çar|Per|Cum|Cmt',
  ],
  [
    'et',
    'jaanuar|veebruar|märts|aprill|mai|juuni|juuli|august|september|oktoober|november|detsember',
    'jaan|veebr|märts|apr|mai|juuni|juuli|aug|sept|okt|nov|dets',
    'pühapäev|esmaspäev|teisipäev|kolmapäev|neljapäev|reede|laupäev',
    'P|E|T|K|N|R|L',
  ],
  [
    'lv',
    'janvāris|februāris|marts|aprīlis|maijs|jūnijs|jūlijs|augusts|septembris|oktobris|novembris|decembris',
    'janv.|febr.|marts|apr.|maijs|jūn.|jūl.|aug.|sept.|okt.|nov.|dec.',
    'Svētdiena|Pirmdiena|Otrdiena|Trešdiena|Ceturtdiena|Piektdiena|Sestdiena',
    'Svētd.|Pirmd.|Otrd.|Trešd.|Ceturtd.|Piektd.|Sestd.',
  ],
  [
    'lt',
    'sausis|vasaris|kovas|balandis|gegužė|birželis|liepa|rugpjūtis|rugsėjis|spalis|lapkritis|gruodis',
    'saus.|vas.|kov.|bal.|geg.|birž.|liep.|rugp.|rugs.|spal.|lapkr.|gruod.',
    'sekmadienis|pirmadienis|antradienis|trečiadienis|ketvirtadienis|penktadienis|šeštadienis',
    'sk|pr|an|tr|kt|pn|št',
    'sausio|vasario|kovo|balandžio|gegužės|birželio|liepos|rugpjūčio|rugsėjo|spalio|lapkričio|gruodžio',
  ],
  [
    'ca',
    'gener|febrer|març|abril|maig|juny|juliol|agost|setembre|octubre|novembre|desembre',
    'gen.|febr.|març|abr.|maig|juny|jul.|ag.|set.|oct.|nov.|des.',
    'diumenge|dilluns|dimarts|dimecres|dijous|divendres|dissabte',
    'dg|dl|dt|dc|dj|dv|ds',
  ],
  [
    'eu',
    'urtarrila|otsaila|martxoa|apirila|maiatza|ekaina|uztaila|abuztua|iraila|urria|azaroa|abendua',
    'urt.|ots.|mar.|api.|mai.|eka.|uzt.|abu.|ira.|urr.|aza.|abe.',
    'igandea|astelehena|asteartea|asteazkena|osteguna|ostirala|larunbata',
    'ig.|al.|ar.|az.|og.|or.|lr.',
  ],
  [
    'gl',
    'xaneiro|febreiro|marzo|abril|maio|xuño|xullo|agosto|setembro|outubro|novembro|decembro',
    'xan.|feb.|mar.|abr.|maio|xuño|xul.|ago.|set.|out.|nov.|dec.',
    'domingo|luns|martes|mércores|xoves|venres|sábado',
    'dom.|luns|mar.|mér.|xov.|ven.|sáb.',
  ],
  [
    'ga',
    'Eanáir|Feabhra|Márta|Aibreán|Bealtaine|Meitheamh|Iúil|Lúnasa|Meán Fómhair|Deireadh Fómhair|Samhain|Nollaig',
    'Ean|Feabh|Márta|Aib|Beal|Meith|Iúil|Lún|MFómh|DFómh|Samh|Noll',
    'Dé Domhnaigh|Dé Luain|Dé Máirt|Dé Céadaoin|Déardaoin|Dé hAoine|Dé Sathairn',
    'Domh|Luan|Máirt|Céad|Déar|Aoine|Sath',
  ],
  [
    'cy',
    'Ionawr|Chwefror|Mawrth|Ebrill|Mai|Mehefin|Gorffennaf|Awst|Medi|Hydref|Tachwedd|Rhagfyr',
    'Ion|Chw|Maw|Ebr|Mai|Meh|Gor|Awst|Medi|Hyd|Tach|Rhag',
    'Dydd Sul|Dydd Llun|Dydd Mawrth|Dydd Mercher|Dydd Iau|Dydd Gwener|Dydd Sadwrn',
    'Sul|Llun|Maw|Mer|Iau|Gwe|Sad',
  ],
  [
    'mt',
    'Jannar|Frar|Marzu|April|Mejju|Ġunju|Lulju|Awwissu|Settembru|Ottubru|Novembru|Diċembru',
    'Jan|Fra|Mar|Apr|Mej|Ġun|Lul|Aww|Set|Ott|Nov|Diċ',
    'Il-Ħadd|It-Tnejn|It-Tlieta|L-Erbgħa|Il-Ħamis|Il-Ġimgħa|Is-Sibt',
    'Ħad|Tne|Tli|Erb|Ħam|Ġim|Sib',
  ],
  [
    'sq',
    'janar|shkurt|mars|prill|maj|qershor|korrik|gusht|shtator|tetor|nëntor|dhjetor',
    'jan|shk|mar|pri|maj|qer|korr|gush|sht|tet|nën|dhj',
    'e diel|e hënë|e martë|e mërkurë|e enjte|e premte|e shtunë',
    'die|hën|mar|mër|enj|pre|sht',
  ],
  [
    'mk',
    'јануари|февруари|март|април|мај|јуни|јули|август|септември|октомври|ноември|декември',
    'јан.|фев.|мар.|апр.|мај|јун.|јул.|авг.|сеп.|окт.|ное.|дек.',
    'недела|понеделник|вторник|среда|четврток|петок|сабота',
    'нед.|пон.|вто.|сре.|чет.|пет.|саб.',
  ],
  [
    'ka',
    'იანვარი|თებერვალი|მარტი|აპრილი|მაისი|ივნისი|ივლისი|აგვისტო|სექტემბერი|ოქტომბერი|ნოემბერი|დეკემბერი',
    'იან|თებ|მარ|აპრ|მაი|ივნ|ივლ|აგვ|სექ|ოქტ|ნოე|დეკ',
    'კვირა|ორშაბათი|სამშაბათი|ოთხშაბათი|ხუთშაბათი|პარასკევი|შაბათი',
    'კვი|ორშ|სამ|ოთხ|ხუთ|პარ|შაბ',
  ],
  [
    'hy',
    'հունվար|փետրվար|մարտ|ապրիլ|մայիս|հունիս|հուլիս|օգոստոս|սեպտեմբեր|հոկտեմբեր|նոյեմբեր|դեկտեմբեր',
    'հնվ|փտվ|մրտ|ապր|մյս|հնս|հլս|օգս|սեպ|հոկ|նոյ|դեկ',
    'կիրակի|երկուշաբթի|երեքշաբթի|չորեքշաբթի|հինգշաբթի|ուրբաթ|շաբաթ',
    'կիր|երկ|երք|չրք|հնգ|ուր|շբթ',
  ],
  [
    'az',
    'yanvar|fevral|mart|aprel|may|iyun|iyul|avqust|sentyabr|oktyabr|noyabr|dekabr',
    'yan|fev|mar|apr|may|iyn|iyl|avq|sen|okt|noy|dek',
    'bazar|bazar ertəsi|çərşənbə axşamı|çərşənbə|cümə axşamı|cümə|şənbə',
    'B.|B.E.|Ç.A.|Ç.|C.A.|C.|Ş.',
  ],
  [
    'kk',
    'Қаңтар|Ақпан|Наурыз|Сәуір|Мамыр|Маусым|Шілде|Тамыз|Қыркүйек|Қазан|Қараша|Желтоқсан',
    'қаң.|ақп.|нау.|сәу.|мам.|мау.|шіл.|там.|қыр.|қаз.|қар.|жел.',
    'жексенбі|дүйсенбі|сейсенбі|сәрсенбі|бейсенбі|жұма|сенбі',
    'жс|дс|сс|ср|бс|жм|сб',
  ],
  [
    'id',
    'Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember',
    'Jan|Feb|Mar|Apr|Mei|Jun|Jul|Agu|Sep|Okt|Nov|Des',
    'Minggu|Senin|Selasa|Rabu|Kamis|Jumat|Sabtu',
    'Min|Sen|Sel|Rab|Kam|Jum|Sab',
  ],
  [
    'ms',
    'Januari|Februari|Mac|April|Mei|Jun|Julai|Ogos|September|Oktober|November|Disember',
    'Jan|Feb|Mac|Apr|Mei|Jun|Jul|Ogo|Sep|Okt|Nov|Dis',
    'Ahad|Isnin|Selasa|Rabu|Khamis|Jumaat|Sabtu',
    'Ahd|Isn|Sel|Rab|Kha|Jum|Sab',
  ],
  [
    'vi',
    'Tháng 1|Tháng 2|Tháng 3|Tháng 4|Tháng 5|Tháng 6|Tháng 7|Tháng 8|Tháng 9|Tháng 10|Tháng 11|Tháng 12',
    'Tháng 1|Tháng 2|Tháng 3|Tháng 4|Tháng 5|Tháng 6|Tháng 7|Tháng 8|Tháng 9|Tháng 10|Tháng 11|Tháng 12',
    'Chủ Nhật|Thứ Hai|Thứ Ba|Thứ Tư|Thứ Năm|Thứ Sáu|Thứ Bảy',
    'CN|Thứ 2|Thứ 3|Thứ 4|Thứ 5|Thứ 6|Thứ 7',
  ],
  [
    'fil',
    'Enero|Pebrero|Marso|Abril|Mayo|Hunyo|Hulyo|Agosto|Setyembre|Oktubre|Nobyembre|Disyembre',
    'Ene|Peb|Mar|Abr|May|Hun|Hul|Ago|Set|Okt|Nob|Dis',
    'Linggo|Lunes|Martes|Miyerkules|Huwebes|Biyernes|Sabado',
    'Lin|Lun|Mar|Miy|Huw|Biy|Sab',
  ],
  [
    'sw',
    'Januari|Februari|Machi|Aprili|Mei|Juni|Julai|Agosti|Septemba|Oktoba|Novemba|Desemba',
    'Jan|Feb|Mac|Apr|Mei|Jun|Jul|Ago|Sep|Okt|Nov|Des',
    'Jumapili|Jumatatu|Jumanne|Jumatano|Alhamisi|Ijumaa|Jumamosi',
    'Jumapili|Jumatatu|Jumanne|Jumatano|Alhamisi|Ijumaa|Jumamosi',
  ],
  [
    'af',
    'Januarie|Februarie|Maart|April|Mei|Junie|Julie|Augustus|September|Oktober|November|Desember',
    'Jan.|Feb.|Mrt.|Apr.|Mei|Jun.|Jul.|Aug.|Sep.|Okt.|Nov.|Des.',
    'Sondag|Maandag|Dinsdag|Woensdag|Donderdag|Vrydag|Saterdag',
    'So.|Ma.|Di.|Wo.|Do.|Vr.|Sa.',
  ],
  [
    'zu',
    'Januwari|Februwari|Mashi|Ephreli|Meyi|Juni|Julayi|Agasti|Septhemba|Okthoba|Novemba|Disemba',
    'Jan|Feb|Mas|Eph|Mey|Jun|Jul|Aga|Sep|Okt|Nov|Dis',
    'ISonto|UMsombuluko|ULwesibili|ULwesithathu|ULwesine|ULwesihlanu|UMgqibelo',
    'Son|Mso|Bil|Tha|Sin|Hla|Mgq',
  ],
  [
    'ja',
    '1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月',
    '1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月',
    '日曜日|月曜日|火曜日|水曜日|木曜日|金曜日|土曜日',
    '日|月|火|水|木|金|土',
  ],
  [
    'zh',
    '一月|二月|三月|四月|五月|六月|七月|八月|九月|十月|十一月|十二月',
    '1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月',
    '星期日|星期一|星期二|星期三|星期四|星期五|星期六',
    '周日|周一|周二|周三|周四|周五|周六',
  ],
  [
    'zh-tw',
    '1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月',
    '1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月',
    '星期日|星期一|星期二|星期三|星期四|星期五|星期六',
    '週日|週一|週二|週三|週四|週五|週六',
  ],
  [
    'ko',
    '1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월',
    '1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월',
    '일요일|월요일|화요일|수요일|목요일|금요일|토요일',
    '일|월|화|수|목|금|토',
  ],
  [
    'hi',
    'जनवरी|फ़रवरी|मार्च|अप्रैल|मई|जून|जुलाई|अगस्त|सितंबर|अक्तूबर|नवंबर|दिसंबर',
    'जन॰|फ़र॰|मार्च|अप्रैल|मई|जून|जुल॰|अग॰|सित॰|अक्तू॰|नव॰|दिस॰',
    'रविवार|सोमवार|मंगलवार|बुधवार|गुरुवार|शुक्रवार|शनिवार',
    'रवि|सोम|मंगल|बुध|गुरु|शुक्र|शनि',
  ],
  [
    'bn',
    'জানুয়ারি|ফেব্রুয়ারি|মার্চ|এপ্রিল|মে|জুন|জুলাই|আগস্ট|সেপ্টেম্বর|অক্টোবর|নভেম্বর|ডিসেম্বর',
    'জানু|ফেব|মার্চ|এপ্রিল|মে|জুন|জুলাই|আগস্ট|সেপ্টেঃ|অক্টোঃ|নভেঃ|ডিসেঃ',
    'রবিবার|সোমবার|মঙ্গলবার|বুধবার|বৃহস্পতিবার|শুক্রবার|শনিবার',
    'রবি|সোম|মঙ্গল|বুধ|বৃহস্পতি|শুক্র|শনি',
  ],
  [
    'ta',
    'ஜனவரி|பிப்ரவரி|மார்ச்|ஏப்ரல்|மே|ஜூன்|ஜூலை|ஆகஸ்ட்|செப்டம்பர்|அக்டோபர்|நவம்பர்|டிசம்பர்',
    'ஜன.|பிப்.|மார்.|ஏப்.|மே|ஜூன்|ஜூலை|ஆக.|செப்.|அக்.|நவ.|டிச.',
    'ஞாயிறு|திங்கள்|செவ்வாய்|புதன்|வியாழன்|வெள்ளி|சனி',
    'ஞாயி.|திங்.|செவ்.|புத.|வியா.|வெள்.|சனி',
  ],
  [
    'ar',
    'كانون الثاني|شباط|آذار|نيسان|أيار|حزيران|تموز|آب|أيلول|تشرين الأول|تشرين الثاني|كانون الأول',
    'كانون الثاني|شباط|آذار|نيسان|أيار|حزيران|تموز|آب|أيلول|تشرين الأول|تشرين الثاني|كانون الأول',
    'الأحد|الإثنين|الثلاثاء|الأربعاء|الخميس|الجمعة|السبت',
    'الأحد|الإثنين|الثلاثاء|الأربعاء|الخميس|الجمعة|السبت',
  ],
  [
    'he',
    'ינואר|פברואר|מרץ|אפריל|מאי|יוני|יולי|אוגוסט|ספטמבר|אוקטובר|נובמבר|דצמבר',
    'ינואר|פברואר|מרץ|אפריל|מאי|יוני|יולי|אוגוסט|ספטמבר|אוקטובר|נובמבר|דצמבר',
    'יום ראשון|יום שני|יום שלישי|יום רביעי|יום חמישי|יום שישי|יום שבת',
    'יום ראשון|יום שני|יום שלישי|יום רביעי|יום חמישי|יום שישי|יום שבת',
  ],
  [
    'fa',
    'ژانويه|فوريه|مارس|آوريل|مه|ژوئن|ژوئيه|اوت|سپتامبر|اُكتبر|نوامبر|دسامبر',
    'ژانويه|فوريه|مارس|آوريل|مه|ژوئن|ژوئيه|اوت|سپتامبر|اُكتبر|نوامبر|دسامبر',
    'يكشنبه|دوشنبه|سه شنبه|چهارشنبه|پنجشنبه|جمعه|شنبه',
    'يكشنبه|دوشنبه|سه شنبه|چهارشنبه|پنجشنبه|جمعه|شنبه',
  ],
  [
    'ur',
    'جنوری|فروری|مارچ|اپریل|مئی|جون|جولائی|اگست|ستمبر|اکتوبر|نومبر|دسمبر',
    'جنوری|فروری|مارچ|اپریل|مئی|جون|جولائی|اگست|ستمبر|اکتوبر|نومبر|دسمبر',
    'اتوار|پير|منگل|بدھ|جمعرات|جمعه|هفته',
    'اتوار|پير|منگل|بدھ|جمعرات|جمعه|هفته',
  ],
];

/** Languages whose months take the after-day form also when the day number follows them. */
const DAY_EITHER_SIDE: ReadonlySet<string> = new Set(['lt']);

/** The text Word writes for day 1 before a month name, by row. */
const FIRST_DAY: ReadonlyMap<string, string> = new Map([
  ['fr', '1er'],
  ['fr-ca', '1er'],
]);

const ROW_OF: ReadonlyMap<string, readonly string[]> = new Map(
  ROWS.map(([tag, ...sections]) => [tag!, sections])
);

const ENGLISH = 'en';

/** A language tag as Word writes one: letters, digits and hyphens, at most 35 characters. */
const LANGUAGE_TAG = /^[A-Za-z0-9-]{1,35}$/;

const namesByRow = new Map<string, DateNames>();

/**
 * The names of a date control's language: the row of its tag, else of the tag with its last
 * parts left off (`sr-Cyrl-ME`, `sr-Cyrl`, `sr`), else English. The tag comes from the file, so
 * it is bounded first, and a lookup in a `Map` matches no prototype key.
 */
export function dateNamesFor(lid: string | undefined): DateNames {
  return namesOfRow(rowKeyOf(lid));
}

function rowKeyOf(lid: string | undefined): string {
  if (lid === undefined || !LANGUAGE_TAG.test(lid)) {
    return ENGLISH;
  }

  let key = lid.toLowerCase();
  while (!ROW_OF.has(key)) {
    const cut = key.lastIndexOf('-');
    if (cut < 0) {
      return ENGLISH;
    }

    key = key.slice(0, cut);
  }

  return key;
}

function namesOfRow(key: string): DateNames {
  const known = namesByRow.get(key);
  if (known) {
    return known;
  }

  const [months, shortMonths, weekdays, shortWeekdays, afterDay] = ROW_OF.get(key)!.map((section) =>
    section.split('|')
  );
  const names: DateNames = {
    months: months!,
    shortMonths: shortMonths!,
    weekdays: weekdays!,
    shortWeekdays: shortWeekdays!,
    monthsAfterDay: afterDay ?? months!,
    dayEitherSide: DAY_EITHER_SIDE.has(key),
    firstDay: FIRST_DAY.get(key),
  };
  namesByRow.set(key, names);
  return names;
}
