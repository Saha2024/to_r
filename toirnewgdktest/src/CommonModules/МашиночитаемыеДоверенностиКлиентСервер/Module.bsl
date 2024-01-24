
#Область СлужебныйПрограммныйИнтерфейс

#Область ПроверкаДоверенности

// Инициализирует структуру для проверки МЧД.
// 
// Возвращаемое значение:
//  Структура:
// * ДанныеДоверенности - ДвоичныеДанные
// * ДанныеПодписи - ДвоичныеДанные
//
Функция НовыеДанныеДляПроверкиМЧД() Экспорт
	
	ДанныеДляПроверки = Новый Структура;
	ДанныеДляПроверки.Вставить("ДанныеДоверенности", Неопределено);
	ДанныеДляПроверки.Вставить("ДанныеПодписи", Неопределено);
	
	Возврат ДанныеДляПроверки;
	
КонецФункции

#КонецОбласти

Функция ЗаголовокБудетОтозвана() Экспорт
	Возврат НСтр("ru = 'Будет отозвана'");
КонецФункции

// Возвращает структуру способов отражения в учете ЭД Машиночитаемая доверенность
// 
// Возвращаемое значение:
//  Структура:
//  * МашиночитаемыеДоверенностиКонтрагентов - Строка
//  * МашиночитаемыеДоверенностиОрганизаций - Строка
Функция СпособыОтраженияВУчете() Экспорт
	
	Способы = Новый Структура;
	Способы.Вставить("МашиночитаемыеДоверенностиКонтрагентов", НСтр("ru = 'Машиночитаемая доверенность контрагента'"));
	Способы.Вставить("МашиночитаемыеДоверенностиОрганизаций", НСтр("ru = 'Машиночитаемая доверенность организации'"));
	Возврат Способы;
	
КонецФункции

// Коллекция видов документов
// 
// Возвращаемое значение:
//  Соответствие Из КлючИЗначение- где:
//   * Ключ - Строка - код вида документа, 
//   * Значение - Строка - наименование вида документа.
//
Функция ВидыДокументовФизическихЛиц() Экспорт
	
	ВидыДокументов = Новый Соответствие();
	
	ВидыДокументов.Вставить("07", НСтр("ru = 'Военный билет'"));
	ВидыДокументов.Вставить("10", НСтр("ru = 'Паспорт иностранного гражданина'"));
	ВидыДокументов.Вставить("11", НСтр("ru = 'Свидетельство о рассмотрении ходатайства о признании лица беженцем на территории Российской Федерации по существу'"));
	ВидыДокументов.Вставить("12", НСтр("ru = 'Вид на жительство в Российской Федерации'"));
	ВидыДокументов.Вставить("13", НСтр("ru = 'Удостоверение беженца'"));
	ВидыДокументов.Вставить("15", НСтр("ru = 'Разрешение на временное проживание в Российской Федерации'"));
	ВидыДокументов.Вставить("19", НСтр("ru = 'Свидетельство о предоставлении временного убежища на территории Российской Федерации'"));
	ВидыДокументов.Вставить("21", НСтр("ru = 'Паспорт гражданина Российской Федерации'"));
	ВидыДокументов.Вставить("24", НСтр("ru = 'Удостоверение личности военнослужащего Российской Федерации'"));
	
	Возврат ВидыДокументов;
	
КонецФункции

// Возвращает СтатусВРеестреФНС МЧД, полученной из реестра
//
// Параметры:
//  СтатусДоверенности - Строка - "CREATED" - дата начала действия не наступила, "ACTIVE" - действует,
//                                "EXPIRED" - истекла, "REVOKED/DECLINED" - отменена (отозвана), "" - не запрашивался
//
// Возвращаемое значение:
//  ПеречислениеСсылка.СтатусыМашиночитаемойДоверенностиВРеестреФНС
//
Функция СтатусВРеестреФНС(СтатусДоверенности) Экспорт
	Если СтатусДоверенности = "ACTIVE" Тогда
		Возврат ПредопределенноеЗначение("Перечисление.СтатусыМашиночитаемойДоверенностиВРеестреФНС.Зарегистрировано");
	ИначеЕсли СтатусДоверенности = "CREATED" Тогда
		Возврат ПредопределенноеЗначение("Перечисление.СтатусыМашиночитаемойДоверенностиВРеестреФНС.ДатаНачалаДействияНеНаступила");
	ИначеЕсли СтатусДоверенности = "EXPIRED" Тогда
		Возврат ПредопределенноеЗначение("Перечисление.СтатусыМашиночитаемойДоверенностиВРеестреФНС.ИстекСрокДействия");
	ИначеЕсли СтатусДоверенности = "REVOKED" ИЛИ СтатусДоверенности = "DECLINED" Тогда
		Возврат ПредопределенноеЗначение("Перечисление.СтатусыМашиночитаемойДоверенностиВРеестреФНС.Отозвано");
	Иначе
		Возврат ПредопределенноеЗначение("Перечисление.СтатусыМашиночитаемойДоверенностиВРеестреФНС.ПустаяСсылка");
	КонецЕсли;
КонецФункции

// Возвращает результат проверки на необходимость наличия доверенности на подпись.
//
// Параметры:
//  ИННДоверителя - Строка - ИНН доверителя
//  СвойстваСубъектаСертификата - Структура - см. КриптографияБЭД.СвойстваСубъектаСертификата
//  СвойстваИздателяСертификата - Структура - см. КриптографияБЭД.СвойстваИздателяСертификата
//
// Возвращаемое значение:
//  Булево - Истина, если для подписи требуется МЧД
//
Функция ТребуетсяМашиночитаемаяДоверенность(ИННДоверителя, СвойстваСубъектаСертификата, СвойстваИздателяСертификата) Экспорт
	
	Если Не ЭтоСертификатФизическогоЛица(СвойстваСубъектаСертификата, СвойстваИздателяСертификата) Тогда
		Возврат Ложь;
	КонецЕсли;
	
	ИННПодписанта = СвойстваСубъектаСертификата.ИНН;
	
	ДоверительИНН = СтроковыеФункцииКлиентСервер.ДополнитьСтроку(ИННДоверителя, 12);
	ПодписантИНН = СтроковыеФункцииКлиентСервер.ДополнитьСтроку(ИННПодписанта, 12);
	
	Возврат ДоверительИНН <> ПодписантИНН;
	
КонецФункции

// Возвращает признак является ли субъект сертификата физическим лицом.
// 
// Параметры:
//  СвойстваСубъекта - См. КриптографияБЭД.СвойстваСубъектаСертификата
//  СвойстваИздателя - См. КриптографияБЭД.СвойстваИздателяСертификата
// 
// Возвращаемое значение:
//  Булево - Это субъект сертификата физическое лицо
Функция ЭтоСертификатФизическогоЛица(СвойстваСубъекта, СвойстваИздателя) Экспорт
	
	Если ЭтоСертификатДолжностногоЛицаКазначейства(СвойстваСубъекта, СвойстваИздателя) Тогда
		Возврат Ложь;
	КонецЕсли;
	
	ИНН = СтроковыеФункцииКлиентСервер.УдалитьПовторяющиесяСимволы(СвойстваСубъекта.ИНН, "0");
	Возврат СтрДлина(ИНН) = 12 И СвойстваСубъекта.ОГРНИП = Неопределено;
	
КонецФункции

// Возвращает признак, что передана ссылка на доверенность.
// 
// Параметры:
//  Ссылка - ЛюбаяСсылка
// 
// Возвращаемое значение:
//  Булево - Истина, если передана ссылка на машиночитаемую доверенность.
Функция ЭтоСсылкаНаДоверенность(Ссылка) Экспорт
	
	Результат = Ложь;
	
	ТипСсылки = ТипЗнч(Ссылка);
	Если ТипСсылки = Тип("СправочникСсылка.МашиночитаемыеДоверенностиОрганизаций")
		ИЛИ ТипСсылки = Тип("СправочникСсылка.МашиночитаемыеДоверенностиКонтрагентов") Тогда
		Результат = Истина;
	КонецЕсли;
	
	Возврат Результат;
	
КонецФункции

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Проверяет, является ли сертификат сертификатом должностного лица казначейства.
// 
// Параметры:
//  СвойстваСубъектаСертификата - см. КриптографияБЭД.СвойстваСубъектаСертификата
//  СвойстваИздателяСертификата - см. КриптографияБЭД.СвойстваИздателяСертификата
// 
// Возвращаемое значение:
//  Булево
Функция ЭтоСертификатДолжностногоЛицаКазначейства(СвойстваСубъектаСертификата, СвойстваИздателяСертификата) Экспорт
	
	ИздательКазначейство = ЭтоИННКазначейства(СвойстваИздателяСертификата.ИНН);
	
	ОрганизацияВСубъектеЗаполнена =  ЗначениеЗаполнено(СвойстваСубъектаСертификата.Организация);
	
	Возврат ИздательКазначейство И ОрганизацияВСубъектеЗаполнена;
	
КонецФункции

// Проверяет, принадлежит ли ИНН казначейству.
// 
// Параметры:
//  ИНН - Строка
// 
// Возвращаемое значение:
//  Булево
Функция ЭтоИННКазначейства(Знач ИНН) Экспорт
	
	ИНН = СтроковыеФункцииКлиентСервер.ДополнитьСтроку(ИНН, 12);
	ИННКазначейства = "007710568760";
	
	Возврат ИНН = ИННКазначейства;
	
КонецФункции


#КонецОбласти