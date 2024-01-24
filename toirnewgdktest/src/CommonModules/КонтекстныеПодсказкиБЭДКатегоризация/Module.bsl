////////////////////////////////////////////////////////////////////////////////
// КонтекстныеПодсказкиБЭДКатегоризация: механизм контекстных подсказок.
//
////////////////////////////////////////////////////////////////////////////////

#Область СлужебныйПрограммныйИнтерфейс

// Возвращает признак соответствия категории общему контексту конфигурации.
//
// Параметры:
//  Категория - ПланВидовХарактеристикСсылка.КатегорииНовостей - категория;
//  Значение - Булево, Строка - значение категории. 
//
// Возвращаемое значение:
//  Булево - признак соответствия категории общему контексту конфигурации.
// 
Функция УдовлетворяетОбщемуКонтексту(Категория, Значение) Экспорт
	
	ФункцияРасчета = ФункцияРасчетаВнеконтекстнойКатегории(Категория); 
	
	Параметры = Новый Структура;
	Параметры.Вставить("Значение", Значение); 
	Параметры.Вставить("Результат", Ложь); 
	
	ОбщегоНазначения.ВыполнитьВБезопасномРежиме("Параметры.Результат = КонтекстныеПодсказкиБЭДКатегоризация." 
													+ ФункцияРасчета + "(Параметры.Значение)", Параметры); 
	
	Возврат Параметры.Результат;
	
КонецФункции

// Возвращает признак принадлежности категории к внеконтекстным.
//
// Параметры:
//  Категория - ПланВидовХарактеристикСсылка.КатегорииНовостей - категория;
//
// Возвращаемое значение:
//  Булево - признак принадлежности категории к внеконтекстным.
// 
Функция ЭтоВнеконтекстнаяКатегория(Категория) Экспорт
	
	СписокВнеконтекстныхКатегорий = КонтекстныеПодсказкиБЭДПовтИсп.СписокВнеконтекстныхКатегорий();
	
	Возврат СписокВнеконтекстныхКатегорий.Получить(Категория) <> Неопределено;
	
КонецФункции

// Возвращает параметры доступных значений для списка категорий.
//
// Параметры:
//  Категории - Массив из ПланВидовХарактеристикСсылка.КатегорииНовостей - категории получения возможного контекста;
//
// Возвращаемое значение:
//  Соответствие - параметры доступных значений категорий.
//    * Ключ - ПланВидовХарактеристикСсылка.КатегорииНовостей - категория;
//    * Значение - Структура - параметры доступных значений категории.
//      * Значения - СписокЗначений - список доступных значений категории.
//      * Типы - ОписаниеТипов - доступные типы значений категорий.
//
Функция ВозможныйКонтекст(Категории) Экспорт

	ВозможныйКонтекст = Новый Соответствие;
	
	Для Каждого Категория Из Категории Цикл 
		
		Если Не ЗначениеЗаполнено(Категория) Тогда
			Продолжить;
		КонецЕсли;
		
		ПараметрыДоступныхЗначений = Новый Структура; 
		
		Параметры = Новый Структура;
		Параметры.Вставить("Результат", Новый Структура);
		
		ТипЗначенияКатегории = Категория.ТипЗначения;
		
		Если ТипЗначенияКатегории.Типы().Найти(Тип("Булево")) = Неопределено Тогда
			
			КодКатегории = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Категория, "Код");
			ОбщегоНазначения.ВыполнитьВБезопасномРежиме("Параметры.Результат = КонтекстныеПодсказкиБЭДКатегоризация.ПараметрыДоступныхЗначенийКатегории_" + КодКатегории + "()", Параметры); 
			ПараметрыДоступныхЗначений = Параметры.Результат;

		Иначе
			
			ПараметрыДоступныхЗначений = ПараметрыДоступныхЗначенийКатегорииЛогическогоТипа(); 
			
		КонецЕсли;
				
		ВозможныйКонтекст.Вставить(Категория, ПараметрыДоступныхЗначений);
		
	КонецЦикла;
	
	Возврат ВозможныйКонтекст;
	
КонецФункции

#Область ВычислениеЗначенийКонтекстныхКатегорий

Функция ЕстьУчетнаяЗаписьЭДО(Организация) Экспорт 
	
	Организации = ОбщегоНазначенияКлиентСервер.ЗначениеВМассиве(Организация);
	Возврат СинхронизацияЭДО.ЕстьУчетныеЗаписи(Организации);
	
КонецФункции

Функция ОператорАбонента(ИдентификаторАбонента) Экспорт
	
	Возврат СинхронизацияЭДО.ОператорАбонента(ИдентификаторАбонента);
	
КонецФункции

Функция КонтрагентПодключенКЭДО(Контрагент) Экспорт
	    
	Возврат ИнтеграцияЭДО.КонтрагентПодключенКСервису1СЭДО(Контрагент)

КонецФункции

Функция СтатусДокументооборота(УчетныйДокумент) Экспорт
	
	Документооборот = ИнтеграцияЭДО.ОсновнойЭлектронныйДокументОбъектаУчета(УчетныйДокумент);
	СостояниеЭДО = ЭлектронныеДокументыЭДО.СостояниеДокумента(Документооборот);
	
	Если Не ЗначениеЗаполнено(СостояниеЭДО) Тогда
		Возврат "";
	КонецЕсли;

	Возврат ОбщегоНазначения.ИмяЗначенияПеречисления(СостояниеЭДО);
	
КонецФункции

Функция ЗначениеТипаЭД(ТипЭД) Экспорт
	
	Если Не ЗначениеЗаполнено(ТипЭД) Тогда
		Возврат "";
	КонецЕсли;
	
	Возврат ОбщегоНазначения.ИмяЗначенияПеречисления(ТипЭД);
	
КонецФункции

Функция ЗначениеСостоянияВерсииЭД(СостояниеВерсииЭД) Экспорт
	
	Если Не ЗначениеЗаполнено(СостояниеВерсииЭД) Тогда
		Возврат "";
	КонецЕсли; 
	
	Возврат ОбщегоНазначения.ИмяЗначенияПеречисления(СостояниеВерсииЭД);
	
КонецФункции

Функция ЗначениеВидаЭД(ВидЭД) Экспорт
	
	Если Не ЗначениеЗаполнено(ВидЭД) Тогда
		Возврат "";
	КонецЕсли;
	 	
	Возврат ОбщегоНазначения.ИмяЗначенияПеречисления(ВидЭД);
	
КонецФункции

Функция ЗначениеНаправленияЭД(НаправлениеЭД) Экспорт
	
	Если Не ЗначениеЗаполнено(НаправлениеЭД) Тогда
		Возврат "";
	КонецЕсли;
	
	Возврат ОбщегоНазначения.ИмяЗначенияПеречисления(НаправлениеЭД);
	
КонецФункции

Функция СуществуютНеверныеПодписиФайла(ЭлектронныйДокумент) Экспорт
	
	Если Не ЗначениеЗаполнено(ЭлектронныйДокумент) Тогда
		Возврат Ложь;
	КонецЕсли;
	
	ЭлектронныеПодписи = ЭлектронныеДокументыЭДО.УстановленныеПодписи(ЭлектронныйДокумент);
	
	Для Каждого СтрокаЭлектроннойПодписи Из ЭлектронныеПодписи Цикл
		Если Не СтрокаЭлектроннойПодписи.ПодписьВерна Тогда
			Возврат Истина;
		КонецЕсли;
	КонецЦикла;
	
	Возврат Ложь;
	
КонецФункции

Функция СвязьОрганизацииСКонтрагентомНастроена(Организация, Контрагент) Экспорт
	
	Возврат СинхронизацияЭДО.СвязьНастроена(Организация, Контрагент);
	
КонецФункции

Функция СтатусКонтроляОтраженияВУЧете(Документооборот) Экспорт	

	Возврат ИнтеграцияЭДО.СтатусКонтроляОтраженияВУЧете(Документооборот);
	
КонецФункции

Функция КодОператораУчетнойЗаписиОрганизации(Организация) Экспорт

	Запросы = Новый Массив;
	Отбор = СинхронизацияЭДО.НовыйОтборУчетныхЗаписей();
	Отбор.Организация = "&Организация";
	ЗапросУчетныхЗаписей = СинхронизацияЭДО.ЗапросУчетныхЗаписей("УчетныеЗаписи", Отбор);
	
	Запросы.Добавить(ЗапросУчетныхЗаписей);
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	|	УчетныеЗаписи.ОператорЭДО
	|ИЗ
	|	УчетныеЗаписи";
	ИтоговыйЗапрос = ОбщегоНазначенияБЭД.СоединитьЗапросы(Запрос, Запросы);
	ИтоговыйЗапрос.УстановитьПараметр("Организация", Организация);
	
	Выборка = ИтоговыйЗапрос.Выполнить().Выбрать();
	
	Если Выборка.Следующий() Тогда
		Возврат Выборка.ОператорЭДО;
	КонецЕсли;

	Возврат "";
КонецФункции

Функция СуществуютСертификатыСИстекающимСрокомДействияДляОрганизации(Организация) Экспорт
	
	Возврат СуществуетДействующийНаДатуСертификат(ДобавитьМесяц(ТекущаяДатаСеанса(), 1), Организация);
	
КонецФункции

Функция СуществуютСертификатыСИстекшимСрокомДействияДляОрганизации(Организация) Экспорт
	
	Возврат СуществуетДействующийНаДатуСертификат(ТекущаяДатаСеанса(), Организация);
	
КонецФункции

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Функция СуществуетДействующийНаДатуСертификат(Дата, Организация = Неопределено)
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ ПЕРВЫЕ 1
		|	СертификатыКлючейЭлектроннойПодписиИШифрования.Ссылка КАК Ссылка
		|ИЗ
		|	Справочник.СертификатыКлючейЭлектроннойПодписиИШифрования КАК СертификатыКлючейЭлектроннойПодписиИШифрования
		|ГДЕ
		|	СертификатыКлючейЭлектроннойПодписиИШифрования.ДействителенДо < &ДействителенДо";
	
	Запрос.Параметры.Вставить("ДействителенДо", Дата);
	
	Если ЗначениеЗаполнено(Организация) Тогда 
		Запрос.Текст = Запрос.Текст + " И СертификатыКлючейЭлектроннойПодписиИШифрования.Организация = &Организация";
		Запрос.Параметры.Вставить("Организация", Организация);
	КонецЕсли;
	
	Возврат Не Запрос.Выполнить().Пустой();

КонецФункции

Функция ПараметрыДоступныхЗначенийКатегорииЛогическогоТипа()
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Булево"));
	
	Значения = Новый СписокЗначений;
	Значения.Добавить(Истина);
	Значения.Добавить(Ложь);
	
	ПараметрыДоступныхЗначений.Вставить("Значения", Значения);
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

Функция ФункцияРасчетаВнеконтекстнойКатегории(Категория)
	
	СписокВнеконтекстныхКатегорий = КонтекстныеПодсказкиБЭДПовтИсп.СписокВнеконтекстныхКатегорий();
	
	Возврат СписокВнеконтекстныхКатегорий.Получить(Категория);
	
КонецФункции

#Область ВычислениеЗначенийВнеконтекстныхКатегорий

Функция СуществуетКонтрагентОператора(КодОператора) Экспорт	
	
	Возврат СинхронизацияЭДО.СуществуетКонтрагентОператора(КодОператора);
	
КонецФункции

Функция СуществуетДокументСоСтатусомОтраженияВУчете(Статус) Экспорт
	
	Возврат ИнтеграцияЭДО.СуществуетДокументСоСтатусомОтраженияВУчете(Статус);
	
КонецФункции

Функция СуществуетНеобработанныйДокументВДеревеДействийДляРаздела(Раздел) Экспорт
	
	Если ДоступныеДействияДереваДействий().Значения.НайтиПоЗначению(Раздел) = Неопределено Тогда 
		Возврат Ложь;	
	КонецЕсли;	
		
	ПараметрыТекущихЭД = ИнтерфейсДокументовЭДО.НовыеПараметрыОпределенияТекущихЭлектронныхДокументов();
	ПараметрыТекущихЭД.Раздел = Раздел;
	ПараметрыТекущихЭД.КоличествоПолучаемыхЗаписей = 1;
		
	Запрос = ИнтерфейсДокументовЭДО.ЗапросКоличестваТекущихЭлектронныхДокументов(ПараметрыТекущихЭД);
	
	Выборка = Запрос.Выполнить().Выбрать();
	
	Если Выборка.Следующий() Тогда
		Возврат Выборка._Счетчик <> 0;
	КонецЕсли;
	
	Возврат Ложь;
	
КонецФункции

Функция ОтсутствуетНеобработанныйДокументВДеревеДействийДляРаздела(Раздел) Экспорт
	
	Возврат Не СуществуетНеобработанныйДокументВДеревеДействийДляРаздела(Раздел);
	
КонецФункции

Функция СоздаватьЭлектронныеПодписиНаСервере(Условие) Экспорт
	
	Возврат ЭлектроннаяПодпись.СоздаватьЭлектронныеПодписиНаСервере() = Условие;

КонецФункции

Функция ПроверятьЭлектронныеПодписиНаСервере(Условие) Экспорт
	
	Возврат ЭлектроннаяПодпись.ПроверятьЭлектронныеПодписиНаСервере() = Условие;
	
КонецФункции

Функция СертификатыЕстьВСписке(Условие) Экспорт
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ ПЕРВЫЕ 1
	|	СертификатыКлючейЭлектроннойПодписиИШифрования.Ссылка КАК Ссылка
	|ИЗ
	|	Справочник.СертификатыКлючейЭлектроннойПодписиИШифрования КАК СертификатыКлючейЭлектроннойПодписиИШифрования
	|ГДЕ
	|	СертификатыКлючейЭлектроннойПодписиИШифрования.Отпечаток <> """"";
	
	Результат = Не Запрос.Выполнить().Пустой();
	
	Возврат Результат = Условие;

КонецФункции

Функция СертификатыЕстьВЛичномСписке(Условие) Экспорт

	ОписанияЗапросов = Новый Массив;

	Отбор = КриптографияБЭД.НовыйОтборСертификатов();
	Отбор.ДоступныеТекущемуПользователю = Истина;

	ОписанияЗапросов.Добавить(КриптографияБЭД.ЗапросДействующихСертификатов("ДоступныеСертификаты", Отбор));
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	ДоступныеСертификаты.Ссылка КАК Сертификат,
		|	ДоступныеСертификаты.Отпечаток КАК Отпечаток
		|ИЗ
		|	ДоступныеСертификаты КАК ДоступныеСертификаты";
	
	Запрос = ОбщегоНазначенияБЭД.СоединитьЗапросы(Запрос, ОписанияЗапросов);
	
	ТаблицаДоступныхСертификатов = Запрос.Выполнить();

	Возврат Не ТаблицаДоступныхСертификатов.Пустой();
	
КонецФункции

Функция СуществуютУчетныеЗаписи(Условие) Экспорт
	
	Возврат СинхронизацияЭДО.ЕстьУчетныеЗаписи() = Условие;

КонецФункции

Функция СуществуютУчетныеЗаписиОператора(КодОператора) Экспорт
	
	Запросы = Новый Массив;
	Отбор = СинхронизацияЭДО.НовыйОтборУчетныхЗаписей();
	Отбор.Оператор = "&Оператор";
	ЗапросУчетныхЗаписей = СинхронизацияЭДО.ЗапросУчетныхЗаписей("УчетныеЗаписи", Отбор);
	
	Запросы.Добавить(ЗапросУчетныхЗаписей);
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ ПЕРВЫЕ 1
	|	УчетныеЗаписи.ИдентификаторЭДО
	|ИЗ
	|	УчетныеЗаписи";
	ИтоговыйЗапрос = ОбщегоНазначенияБЭД.СоединитьЗапросы(Запрос, Запросы);
	ИтоговыйЗапрос.УстановитьПараметр("Оператор", КодОператора);
	
	Возврат Не ИтоговыйЗапрос.Выполнить().Пустой();
	
КонецФункции

Функция СуществуютЭлементыВСправочнике(ИмяСправочника) Экспорт

	ОбъектМетаданных = Метаданные.Справочники.Найти(ИмяСправочника);
	
	Если ОбъектМетаданных = Неопределено Или Не ПравоДоступа("Чтение", ОбъектМетаданных) Тогда
		Возврат Ложь;
	КонецЕсли;
		
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ ПЕРВЫЕ 1
		|	Данные.Ссылка КАК Ссылка
		|ИЗ
		|	Справочник.%1 КАК Данные";
	
	Запрос.Текст = СтрШаблон(Запрос.Текст, ИмяСправочника);

	Результат = Не Запрос.Выполнить().Пустой();	
	
	Возврат Результат;
	
КонецФункции

Функция ОтсутствуютЭлементыВСправочнике(ИмяСправочника) Экспорт
	
	Возврат Не СуществуютЭлементыВСправочнике(ИмяСправочника);
	
КонецФункции

Функция СуществуютСертификатыСИстекающимСрокомДействия(Условие) Экспорт
	
	Возврат СуществуетДействующийНаДатуСертификат(ДобавитьМесяц(ТекущаяДатаСеанса(), 1)); 
	
КонецФункции

Функция СуществуютСертификатыСИстекшимСрокомДействия(Условие) Экспорт
	
	Возврат СуществуетДействующийНаДатуСертификат(ТекущаяДатаСеанса()); 
	
КонецФункции

#КонецОбласти

#Область ВычислениеПараметровДоступныхЗначенийКатегорий

// Состояние документооборота
Функция ПараметрыДоступныхЗначенийКатегории_LED_DocSts() Экспорт
	
	СписокДоступныхЗначений = Новый СписокЗначений;
	
	Для Каждого СостояниеЭДО Из Перечисления.СостоянияДокументовЭДО Цикл
		Имя = ОбщегоНазначения.ИмяЗначенияПеречисления(СостояниеЭДО);
		Синоним = Строка(СостояниеЭДО);
		СписокДоступныхЗначений.Добавить(Имя, Синоним);
	КонецЦикла;
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", СписокДоступныхЗначений); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Тип электронного документа
Функция ПараметрыДоступныхЗначенийКатегории_LED_EDType() Экспорт
	
	СписокДоступныхЗначений = Новый СписокЗначений;
	
	Для Каждого ТипДокумента Из Перечисления.ТипыДокументовЭДО Цикл
		Имя = ОбщегоНазначения.ИмяЗначенияПеречисления(ТипДокумента);
		Синоним = Строка(ТипДокумента);
		СписокДоступныхЗначений.Добавить(Имя, Синоним);
	КонецЦикла;
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", СписокДоступныхЗначений); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Направление электронного документа
Функция ПараметрыДоступныхЗначенийКатегории_LED_DocDirection() Экспорт
	
	СписокДоступныхЗначений = Новый СписокЗначений;
	
	Для Каждого НаправленияЭДО Из Перечисления.НаправленияЭДО Цикл
		Имя = ОбщегоНазначения.ИмяЗначенияПеречисления(НаправленияЭДО);
		Синоним = Строка(НаправленияЭДО);
		СписокДоступныхЗначений.Добавить(Имя, Синоним);
	КонецЦикла;
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", СписокДоступныхЗначений); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Код оператора учетной записи организации
Функция ПараметрыДоступныхЗначенийКатегории_LED_OrgLoginOfOper() Экспорт
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Неопределено); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Код оператора контрагента
Функция ПараметрыДоступныхЗначенийКатегории_LED_CustomerOperCode() Экспорт
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Неопределено); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Существует контрагент оператора X
Функция ПараметрыДоступныхЗначенийКатегории_LED_ExistCustOfOper() Экспорт
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Неопределено); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Существует учетные записи оператора X
Функция ПараметрыДоступныхЗначенийКатегории_LED_AccOfOperIsExist() Экспорт
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Неопределено); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Существует ЭД со статусом отражения в учете X
Функция ПараметрыДоступныхЗначенийКатегории_LED_DocStsOfAccEx() Экспорт
	
	Возврат ПараметрыДоступныхЗначенийКатегории_LED_DocAccSts();
	
КонецФункции

// Существует не обработанные ЭД для действия X в дереве действий
Функция ПараметрыДоступныхЗначенийКатегории_LED_NotPrcDocInActTr() Экспорт
	
	Возврат ДоступныеДействияДереваДействий();
	
КонецФункции

// Отсутствуют не обработанные ЭД для действия X в дереве действий
Функция ПараметрыДоступныхЗначенийКатегории_LED_EmpPrcDocInActTr() Экспорт
	
	Возврат ДоступныеДействияДереваДействий();
	
КонецФункции

// Статусы отражения в учете документов
Функция ПараметрыДоступныхЗначенийКатегории_LED_DocAccSts() Экспорт
	
	Состояния = ИнтеграцияЭДО.СостоянияКонтроляОтраженияВУчете();
	
	Значения = Новый СписокЗначений;
	Значения.Добавить(Состояния.СопоставитьНоменклатуру, НСтр("ru='Сопоставить номенклатуру'"));
	Значения.Добавить(Состояния.СоздатьУчетныйДокумент,  НСтр("ru='Создать учетный документ'"));
	Значения.Добавить(Состояния.ПровестиУчетныйДокумент, НСтр("ru='Провести учетный документ'"));
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Значения); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

// Существуют элементы в справочнике Х
Функция ПараметрыДоступныхЗначенийКатегории_LED_ElemExistInCat() Экспорт
	
	Возврат ПараметрыДоступныхЗначенийДанныеСправочников();
	
КонецФункции

// Отсутствуют элементы в справочнике Х
Функция ПараметрыДоступныхЗначенийКатегории_LED_ElemNotExInCat() Экспорт
	
	Возврат ПараметрыДоступныхЗначенийДанныеСправочников();
	
КонецФункции

Функция ПараметрыДоступныхЗначенийДанныеСправочников()
	
	Значения = Новый СписокЗначений;
	
	Для Каждого Спр Из Метаданные.Справочники Цикл
		Значения.Добавить(Спр.Имя, Спр.Синоним);
	КонецЦикла;
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Значения); 
	
	Возврат ПараметрыДоступныхЗначений;
	
КонецФункции

Функция ДоступныеДействияДереваДействий()
	
	ОписаниеРазделов = Новый ТаблицаЗначений;
	ОписаниеРазделов.Колонки.Добавить("Имя");
	ОписаниеРазделов.Колонки.Добавить("Представление");
	ОписаниеРазделов.Колонки.Добавить("Видимость");
	ОписаниеРазделов.Колонки.Добавить("РассчитыватьКоличество");
	ОписаниеРазделов.Колонки.Добавить("КоличествоРассчитано");

	ИнтерфейсДокументовЭДО.ИнициализироватьРазделыТекущихДел(ОписаниеРазделов);
	
	Значения = Новый СписокЗначений;
	
	Для Каждого Раздел Из ОписаниеРазделов Цикл
		Значения.Добавить(Раздел.Имя, Раздел.Представление);
	КонецЦикла;
	
	ПараметрыДоступныхЗначений = Новый Структура;
	ПараметрыДоступныхЗначений.Вставить("Типы", Новый ОписаниеТипов("Строка"));
	ПараметрыДоступныхЗначений.Вставить("Значения", Значения); 
	
	Возврат ПараметрыДоступныхЗначений;

КонецФункции

#КонецОбласти

#КонецОбласти
