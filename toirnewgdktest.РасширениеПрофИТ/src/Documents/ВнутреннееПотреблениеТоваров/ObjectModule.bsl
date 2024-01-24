
#Область ОбработчикиСобытий

&ИзменениеИКонтроль("ОбработкаПроверкиЗаполнения")
Процедура проф_ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)

	Перем МассивВсехРеквизитов;
	Перем МассивРеквизитовОперации;


	МассивНепроверяемыхРеквизитов = Новый Массив;

	// Проверка количества в т.ч. товары.
	ОбщегоНазначенияУТ.ПроверитьЗаполнениеКоличества(ЭтотОбъект, ПроверяемыеРеквизиты, Отказ);

	// Проверка характеристик в т.ч. товары.
	ФОИспользоватьХарактеристикиНоменклатуры = ПолучитьФункциональнуюОпцию("торо_ИспользоватьХарактеристикиНоменклатуры");
	МассивНепроверяемыхРеквизитов.Добавить("Товары.Характеристика");
	Если ФОИспользоватьХарактеристикиНоменклатуры Тогда
		НоменклатураСервер.ПроверитьЗаполнениеХарактеристик(ЭтотОбъект,МассивНепроверяемыхРеквизитов,Отказ);
	КонецЕсли;

	// Если накладная по заказу - то код строки должен быть заполнен.
	Если Не ЗначениеЗаполнено(ЗаказНаВнутреннееПотребление) Тогда
		МассивНепроверяемыхРеквизитов.Добавить("Товары.КодСтроки");
	КонецЕсли;

	Если ЕстьРасхождения Тогда
		МассивНепроверяемыхРеквизитов.Добавить("Товары");
	КонецЕсли;

	ФОИспользоватьСерииНоменклатуры = ПолучитьФункциональнуюОпцию("ИспользоватьСерииНоменклатуры");
	МассивНепроверяемыхРеквизитов.Добавить("Товары.Серия");
	Если ФОИспользоватьСерииНоменклатуры = Истина тогда
		НоменклатураСервер.ПроверитьЗаполнениеСерий(ЭтотОбъект,
		Отказ,
		МассивНепроверяемыхРеквизитов);
	КонецЕсли;
	#Вставка
	//++ Проф-ИТ, #183, Соловьев А.А., 03.10.2023
	Если проф_ХозяйственнаяОперация <> Перечисления.проф_ХозяйственнаяОперация.ПередачаВЭксплуатацию Тогда 
		МассивНепроверяемыхРеквизитов.Добавить("Товары.проф_КатегорииЭксплуатации");
		МассивНепроверяемыхРеквизитов.Добавить("Товары.проф_ФизическоеЛицо");
	Иначе
		//++ Проф-ИТ, #303, Корнилов М.С., 13.10.2023
		проф_ПроверитьТЧТоварыИнвентарныйНомер(Отказ);
		//-- Проф-ИТ, #303, Корнилов М.С., 13.10.2023
	КонецЕсли;
	//-- Проф-ИТ, #183, Соловьев А.А., 03.10.2023
	#КонецВставки

	ОбщегоНазначения.УдалитьНепроверяемыеРеквизитыИзМассива(ПроверяемыеРеквизиты,МассивНепроверяемыхРеквизитов);

КонецПроцедуры

&ИзменениеИКонтроль("ОбработкаПроведения")
Процедура проф_ОбработкаПроведения(Отказ, РежимПроведения)

	ФОИспользоватьСерии = Константы.ИспользоватьСерииНоменклатуры.Получить();
	ФОИспользоватьХарактеристики = Константы.торо_ИспользоватьХарактеристикиНоменклатуры.Получить();	
	ИспользуетсяКонтрольОстатков = Константы.торо_ИспользоватьКонтрольОтрицательныхОстатков.Получить();

	Движения.ТоварыНаСкладах.Записывать = Истина;
	Движения.торо_ТоварыНаРуках.Записывать = Истина;
	Движения.торо_РезервыНаРуках.Записывать = Истина;
	Движения.торо_РезервыНаСкладах.Записывать = Истина;
	Движения.ЗаказыНаВнутреннееПотребление.Записывать = Истина;

	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	Движения.проф_ТоварыНаСкладах.Записывать = Истина;	
	//-- Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	#КонецВставки
	// контроль остатков

	Запрос = Новый Запрос;
	Запрос.МенеджерВременныхТаблиц = Новый МенеджерВременныхТаблиц;
	Запрос.Текст = "ВЫБРАТЬ
	|	ВнутреннееПотреблениеТовары.Номенклатура КАК Номенклатура,
	|	&Характеристика КАК Характеристика,
	|	ВнутреннееПотреблениеТовары.Упаковка КАК Упаковка,
	|	ВнутреннееПотреблениеТовары.КоличествоУпаковок КАК КоличествоУпаковок,
	|	ВнутреннееПотреблениеТовары.Количество КАК Количество,
	|	ВнутреннееПотреблениеТовары.КодСтроки КАК КодСтроки,
	|	ВнутреннееПотреблениеТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	ВнутреннееПотреблениеТовары.НомерСтроки КАК НомерСтроки,
	|	&Склад КАК Склад,
	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	|	ВнутреннееПотреблениеТовары.проф_Назначение КАК Назначение,
	//-- Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	#КонецВставки
	|	&Серия КАК Серия
	|ПОМЕСТИТЬ Таблица
	|ИЗ
	|	Документ.ВнутреннееПотреблениеТоваров.Товары КАК ВнутреннееПотреблениеТовары
	|ГДЕ
	|	ВнутреннееПотреблениеТовары.Ссылка = &Ссылка";

	Запрос.УстановитьПараметр("Ссылка",	Ссылка);
	Запрос.УстановитьПараметр("Склад" ,	Склад);

	Если ФОИспользоватьСерии Тогда 
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&Серия", "ВнутреннееПотреблениеТовары.Серия");	
	Иначе
		Запрос.УстановитьПараметр("Серия", Справочники.СерииНоменклатуры.ПустаяСсылка());	
	КонецЕсли;

	Если ФОИспользоватьХарактеристики Тогда 
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&Характеристика", "ВнутреннееПотреблениеТовары.Характеристика");
	Иначе
		Запрос.УстановитьПараметр("Характеристика", Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка());
	КонецЕсли;

	Запрос.Выполнить();

	УстановитьБлокировку();

	// проверка остатков
	Запрос.Текст = "ВЫБРАТЬ
	|	ТабТовары.Номенклатура КАК Номенклатура,
	|	ТабТовары.Склад КАК Склад,
	|	ТабТовары.Характеристика КАК Характеристика,
	|	ТабТовары.Серия КАК Серия,
	|	ТабТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	СУММА(ТабТовары.Количество) КАК Количество,
	|	СпрНоменклатура.ЕдиницаИзмерения КАК ЕдиницаИзмерения
	|ПОМЕСТИТЬ ТЧБезДублей
	|ИЗ
	|	Таблица КАК ТабТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.Номенклатура КАК СпрНоменклатура
	|		ПО ТабТовары.Номенклатура = СпрНоменклатура.Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	ТабТовары.Номенклатура,
	|	ТабТовары.Характеристика,
	|	ТабТовары.Серия,
	|	ТабТовары.Склад,
	|	ТабТовары.ЗаказНаВнутреннееПотребление,
	|	СпрНоменклатура.ЕдиницаИзмерения
	|;
	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ТабТовары.Номенклатура КАК Номенклатура,
	|	ТабТовары.Склад КАК Склад,
	|	ТабТовары.Назначение КАК Назначение,
	|	ТабТовары.Характеристика КАК Характеристика,
	|	ТабТовары.Серия КАК Серия,
	|	ТабТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	СУММА(ТабТовары.Количество) КАК Количество,
	|	СпрНоменклатура.ЕдиницаИзмерения КАК ЕдиницаИзмерения
	|ПОМЕСТИТЬ проф_ТЧБезДублей
	|ИЗ
	|	Таблица КАК ТабТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.Номенклатура КАК СпрНоменклатура
	|		ПО ТабТовары.Номенклатура = СпрНоменклатура.Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	ТабТовары.Номенклатура,
	|	ТабТовары.Характеристика,
	|	ТабТовары.Серия,
	|	ТабТовары.Склад,
	|	ТабТовары.Назначение,
	|	ТабТовары.ЗаказНаВнутреннееПотребление,
	|	СпрНоменклатура.ЕдиницаИзмерения
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ТоварыНаСкладахОстатки.Номенклатура КАК Номенклатура,
	|	ТоварыНаСкладахОстатки.Характеристика КАК Характеристика,
	|	ТоварыНаСкладахОстатки.Склад КАК Склад,
	|	ТоварыНаСкладахОстатки.Назначение КАК Назначение,
	|	ТоварыНаСкладахОстатки.Серия КАК Серия,
	|	ТоварыНаСкладахОстатки.ВНаличииОстаток КАК ВНаличииОстаток
	|ПОМЕСТИТЬ проф_ТоварыНаСкладахОстатки
	|ИЗ
	|	РегистрНакопления.проф_ТоварыНаСкладах.Остатки(
	|			&Дата,
	|			(Номенклатура, Характеристика, Склад, Серия, Назначение) В
	|				(ВЫБРАТЬ
	|					Таблица.Номенклатура,
	|					Таблица.Характеристика,
	|					Таблица.Склад,
	|					Таблица.Серия,
	|					Таблица.Назначение
	|				ИЗ
	|					Таблица КАК Таблица)) КАК ТоварыНаСкладахОстатки
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад,
	|	Серия,
	|	Назначение
	|;
	//-- Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	#КонецВставки
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ТоварыНаСкладахОстатки.Номенклатура КАК Номенклатура,
	|	ТоварыНаСкладахОстатки.Характеристика КАК Характеристика,
	|	ТоварыНаСкладахОстатки.Склад КАК Склад,
	|	ТоварыНаСкладахОстатки.Серия КАК Серия,
	|	ТоварыНаСкладахОстатки.ВНаличииОстаток КАК ВНаличииОстаток
	|ПОМЕСТИТЬ ТоварыНаСкладахОстатки
	|ИЗ
	|	РегистрНакопления.ТоварыНаСкладах.Остатки(
	|			&Дата,
	|			(Номенклатура, Характеристика, Серия, Склад) В
	|				(ВЫБРАТЬ
	|					Таблица.Номенклатура,
	|					Таблица.Характеристика,
	|					Таблица.Серия,
	|					Таблица.Склад
	|				ИЗ
	|					Таблица КАК Таблица)) КАК ТоварыНаСкладахОстатки
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад,
	|	Серия
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_РезервыНаСкладахОстатки.Склад КАК Склад,
	|	торо_РезервыНаСкладахОстатки.Номенклатура КАК Номенклатура,
	|	торо_РезервыНаСкладахОстатки.Характеристика КАК Характеристика,
	|	торо_РезервыНаСкладахОстатки.Серия КАК Серия,
	|	&ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	СУММА(торо_РезервыНаСкладахОстатки.КоличествоОстаток) КАК КоличествоОстаток
	|ПОМЕСТИТЬ РезервыНаСкладах
	|ИЗ
	|	РегистрНакопления.торо_РезервыНаСкладах.Остатки(
	|			&Дата,
	|			(Номенклатура, Характеристика, Серия, Склад) В
	|				(ВЫБРАТЬ
	|					Таблица.Номенклатура,
	|					Таблица.Характеристика,
	|					Таблица.Серия,
	|					Таблица.Склад
	|				ИЗ
	|					Таблица)) КАК торо_РезервыНаСкладахОстатки
	|
	|СГРУППИРОВАТЬ ПО
	|	&ЗаказНаВнутреннееПотребление,
	|	торо_РезервыНаСкладахОстатки.Номенклатура,
	|	торо_РезервыНаСкладахОстатки.Характеристика,
	|	торо_РезервыНаСкладахОстатки.Склад,
	|	торо_РезервыНаСкладахОстатки.Серия
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад,
	|	Серия,
	|	ЗаказНаВнутреннееПотребление
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_ТоварыНаРукахОстатки.Номенклатура КАК Номенклатура,
	|	торо_ТоварыНаРукахОстатки.Характеристика КАК Характеристика,
	|	торо_ТоварыНаРукахОстатки.Склад КАК Склад,
	|	торо_ТоварыНаРукахОстатки.Серия КАК Серия,
	|	торо_ТоварыНаРукахОстатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	торо_ТоварыНаРукахОстатки.КоличествоОстаток КАК КоличествоОстаток
	|ПОМЕСТИТЬ ТоварыНаРуках
	|ИЗ
	|	РегистрНакопления.торо_ТоварыНаРуках.Остатки(
	|			&Дата,
	|			(Номенклатура, Характеристика, Серия, Склад, ЗаказНаВнутреннееПотребление) В
	|				(ВЫБРАТЬ
	|					Таблица.Номенклатура,
	|					Таблица.Характеристика,
	|					Таблица.Серия,
	|					Таблица.Склад,
	|					Таблица.ЗаказНаВнутреннееПотребление
	|				ИЗ
	|					Таблица)) КАК торо_ТоварыНаРукахОстатки
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад,
	|	Серия,
	|	ЗаказНаВнутреннееПотребление
	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ТабТовары.Номенклатура КАК Номенклатура,
	|	ТабТовары.Характеристика КАК Характеристика,
	|	ТабТовары.Склад КАК Склад,
	|	ТабТовары.Серия КАК Серия,
	|	ТабТовары.Назначение КАК Назначение,
	|	ЕСТЬNULL(проф_ТоварыНаСкладахОстатки.ВНаличииОстаток, 0) КАК СвобОстаток,
	|	ТабТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	ТабТовары.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
	|	ЕСТЬNULL(проф_ТоварыНаСкладахОстатки.ВНаличииОстаток, 0) КАК НаСкладе,
	|	ТабТовары.Количество КАК Количество
	|ПОМЕСТИТЬ проф_Остатки
	|ИЗ
	|	проф_ТЧБезДублей КАК ТабТовары
	|		ЛЕВОЕ СОЕДИНЕНИЕ проф_ТоварыНаСкладахОстатки КАК проф_ТоварыНаСкладахОстатки
	|		ПО ТабТовары.Номенклатура = проф_ТоварыНаСкладахОстатки.Номенклатура
	|			И ТабТовары.Характеристика = проф_ТоварыНаСкладахОстатки.Характеристика
	|			И ТабТовары.Серия = проф_ТоварыНаСкладахОстатки.Серия
	|			И ТабТовары.Склад = проф_ТоварыНаСкладахОстатки.Склад
	|			И ТабТовары.Назначение = проф_ТоварыНаСкладахОстатки.Назначение
	|
	|СГРУППИРОВАТЬ ПО
	|	ТабТовары.Номенклатура,
	|	ТабТовары.Характеристика,
	|	ТабТовары.Склад,
	|	ТабТовары.Назначение,
	|	ТабТовары.Серия,
	|	ТабТовары.ЗаказНаВнутреннееПотребление,
	|	проф_ТоварыНаСкладахОстатки.ВНаличииОстаток,
	|	ТабТовары.Количество,
	|	ТабТовары.ЕдиницаИзмерения
	//-- Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	#КонецВставки
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ТабТовары.Номенклатура КАК Номенклатура,
	|	ТабТовары.Характеристика КАК Характеристика,
	|	ТабТовары.Склад КАК Склад,
	|	ТабТовары.Серия КАК Серия,
	|	ЕСТЬNULL(ТоварыНаСкладахОстатки.ВНаличииОстаток, 0) + ЕСТЬNULL(ТоварыНаРуках.КоличествоОстаток, 0) - СУММА(ЕСТЬNULL(Резервы.КоличествоОстаток, 0)) КАК СвобОстаток,
	|	ТабТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	ТабТовары.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
	|	ЕСТЬNULL(ТоварыНаСкладахОстатки.ВНаличииОстаток, 0) - СУММА(ЕСТЬNULL(Резервы.КоличествоОстаток, 0)) КАК НаСкладе,
	|	ЕСТЬNULL(ТоварыНаРуках.КоличествоОстаток, 0) КАК НаРуках,
	|	ТабТовары.Количество КАК Количество
	|ПОМЕСТИТЬ Остатки
	|ИЗ
	|	ТЧБезДублей КАК ТабТовары
	|		ЛЕВОЕ СОЕДИНЕНИЕ ТоварыНаСкладахОстатки КАК ТоварыНаСкладахОстатки
	|		ПО ТабТовары.Номенклатура = ТоварыНаСкладахОстатки.Номенклатура
	|			И ТабТовары.Характеристика = ТоварыНаСкладахОстатки.Характеристика
	|			И ТабТовары.Серия = ТоварыНаСкладахОстатки.Серия
	|			И ТабТовары.Склад = ТоварыНаСкладахОстатки.Склад
	|		ЛЕВОЕ СОЕДИНЕНИЕ РезервыНаСкладах КАК Резервы
	|		ПО ТабТовары.Склад = Резервы.Склад
	|			И ТабТовары.Номенклатура = Резервы.Номенклатура
	|			И (&ВнутреннийЗаказ)
	|			И ТабТовары.Характеристика = Резервы.Характеристика
	|			И ТабТовары.Серия = Резервы.Серия
	|		ЛЕВОЕ СОЕДИНЕНИЕ ТоварыНаРуках КАК ТоварыНаРуках
	|		ПО ТабТовары.Склад = ТоварыНаРуках.Склад
	|			И ТабТовары.Номенклатура = ТоварыНаРуках.Номенклатура
	|			И ТабТовары.ЗаказНаВнутреннееПотребление = ТоварыНаРуках.ЗаказНаВнутреннееПотребление
	|			И ТабТовары.Характеристика = ТоварыНаРуках.Характеристика
	|			И ТабТовары.Серия = ТоварыНаРуках.Серия
	|
	|СГРУППИРОВАТЬ ПО
	|	ТабТовары.Номенклатура,
	|	ТабТовары.Характеристика,
	|	ТабТовары.Склад,
	|	ТабТовары.Серия,
	|	ТабТовары.ЗаказНаВнутреннееПотребление,
	|	ТоварыНаСкладахОстатки.ВНаличииОстаток,
	|	ТабТовары.Количество,
	|	ТабТовары.ЕдиницаИзмерения,
	|	ТоварыНаРуках.КоличествоОстаток";


	Если ПотреблениеПоЗаказам Тогда 
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&ВнутреннийЗаказ", "ТабТовары.ЗаказНаВнутреннееПотребление <> Резервы.ЗаказНаВнутреннееПотребление");
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&ЗаказНаВнутреннееПотребление", "торо_РезервыНаСкладахОстатки.ЗаказНаВнутреннееПотребление");
	Иначе
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&ЗаказНаВнутреннееПотребление", "ЗНАЧЕНИЕ(Документ.ЗаказНаВнутреннееПотребление.ПустаяСсылка)");
		Запрос.УстановитьПараметр("ВнутреннийЗаказ", Истина);
	КонецЕсли;
	Запрос.УстановитьПараметр("Дата", Новый Граница(МоментВремени(), ВидГраницы.Включая));

	Запрос.Выполнить();
	ПроверитьПревышениеОстатка(Запрос.МенеджерВременныхТаблиц, Отказ, ФОИспользоватьСерии, ФОИспользоватьХарактеристики);

	Если Отказ И ИспользуетсяКонтрольОстатков Тогда 
		Возврат;
	ИначеЕсли НЕ ИспользуетсяКонтрольОстатков Тогда  
		Отказ = Ложь;
	КонецЕсли;

	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 23.08.2023
	проф_ОбщегоНазначенияВызовСервера.ПроверитьПревышениеОстатка(Запрос.МенеджерВременныхТаблиц, Отказ, ФОИспользоватьСерии, ФОИспользоватьХарактеристики, ЭтотОбъект, Склад);
	Если Отказ И ИспользуетсяКонтрольОстатков Тогда 
		Возврат;
	ИначеЕсли НЕ ИспользуетсяКонтрольОстатков Тогда  
		Отказ = Ложь;
	КонецЕсли;
	//-- Проф-ИТ, #27, Соловьев А.А., 23.08.2023
	#КонецВставки
	Если ПотреблениеПоЗаказам Тогда
		СписатьЗаказыНаВнутреннееПотребление(Запрос.МенеджерВременныхТаблиц, Отказ, ФОИспользоватьСерии, ФОИспользоватьХарактеристики);
		Если Отказ Тогда 
			Возврат;
		КонецЕсли;
		Запрос.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_РезервыНаРукахОстатки.Склад КАК Склад,
		|	торо_РезервыНаРукахОстатки.Номенклатура КАК Номенклатура,
		|	торо_РезервыНаРукахОстатки.Характеристика КАК Характеристика,
		|	торо_РезервыНаРукахОстатки.Серия КАК Серия,
		|	торо_РезервыНаРукахОстатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	торо_РезервыНаРукахОстатки.ДокументРезерв КАК ДокументРезерв,
		|	торо_РезервыНаРукахОстатки.КоличествоОстаток КАК КоличествоОстаток
		|ПОМЕСТИТЬ РезервыНаРуках
		|ИЗ
		|	РегистрНакопления.торо_РезервыНаРуках.Остатки(
		|			&Дата,
		|			(ЗаказНаВнутреннееПотребление, Номенклатура, Характеристика, Серия, Склад) В
		|				(ВЫБРАТЬ
		|					Таблица.ЗаказНаВнутреннееПотребление,
		|					Таблица.Номенклатура,
		|					Таблица.Характеристика,
		|					Таблица.Серия,
		|					Таблица.Склад
		|				ИЗ
		|					Таблица)) КАК торо_РезервыНаРукахОстатки
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	Характеристика,
		|	Номенклатура,
		|	Склад,
		|	Серия,
		|	ЗаказНаВнутреннееПотребление
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_РезервыНаСкладахОстатки.Склад КАК Склад,
		|	торо_РезервыНаСкладахОстатки.Номенклатура КАК Номенклатура,
		|	торо_РезервыНаСкладахОстатки.Характеристика КАК Характеристика,
		|	торо_РезервыНаСкладахОстатки.Серия КАК Серия,
		|	торо_РезервыНаСкладахОстатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	торо_РезервыНаСкладахОстатки.КоличествоОстаток КАК КоличествоОстаток,
		|	торо_РезервыНаСкладахОстатки.ДокументРезерв КАК ДокументРезерв
		|ПОМЕСТИТЬ РезервыНаСкладахСДокументами
		|ИЗ
		|	РегистрНакопления.торо_РезервыНаСкладах.Остатки(
		|			&Дата,
		|			(ЗаказНаВнутреннееПотребление, Номенклатура, Характеристика, Серия, Склад) В
		|				(ВЫБРАТЬ
		|					Таблица.ЗаказНаВнутреннееПотребление,
		|					Таблица.Номенклатура,
		|					Таблица.Характеристика,
		|					Таблица.Серия,
		|					Таблица.Склад
		|				ИЗ
		|					Таблица)) КАК торо_РезервыНаСкладахОстатки
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	Номенклатура,
		|	Характеристика,
		|	Склад,
		|	Серия,
		|	ЗаказНаВнутреннееПотребление
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Остатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	Остатки.Склад КАК Склад,
		|	Остатки.Номенклатура КАК Номенклатура,
		|	Остатки.Характеристика КАК Характеристика,
		|	Остатки.Количество КАК Количество,
		|	Остатки.Серия КАК Серия,
		|	Остатки.НаСкладе КАК НаСкладе,
		|	Остатки.НаРуках КАК НаРуках
		|ИЗ
		|	Остатки КАК Остатки
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Остатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	Остатки.Склад КАК Склад,
		|	Остатки.Номенклатура КАК Номенклатура,
		|	Остатки.Характеристика КАК Характеристика,
		|	Остатки.Количество КАК Количество,
		|	Остатки.Серия КАК Серия,
		|	Остатки.НаСкладе КАК НаСкладе,
		|	Остатки.НаРуках КАК НаРуках,
		|	ЕСТЬNULL(РезервыНаСкладах.ДокументРезерв, ЗНАЧЕНИЕ(Документ.торо_РезервПодВнутреннийЗаказ.ПустаяСсылка)) КАК ДокументРезерв,
		|	ЕСТЬNULL(РезервыНаСкладах.КоличествоОстаток, 0) КАК РезервНаСкладе
		|ИЗ
		|	Остатки КАК Остатки
		|		ЛЕВОЕ СОЕДИНЕНИЕ РезервыНаСкладахСДокументами КАК РезервыНаСкладах
		|		ПО Остатки.Склад = РезервыНаСкладах.Склад
		|			И Остатки.Номенклатура = РезервыНаСкладах.Номенклатура
		|			И Остатки.Характеристика = РезервыНаСкладах.Характеристика
		|			И Остатки.Серия = РезервыНаСкладах.Серия
		|			И Остатки.ЗаказНаВнутреннееПотребление = РезервыНаСкладах.ЗаказНаВнутреннееПотребление
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Остатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	Остатки.Склад КАК Склад,
		|	Остатки.Номенклатура КАК Номенклатура,
		|	Остатки.Характеристика КАК Характеристика,
		|	Остатки.Количество КАК Количество,
		|	Остатки.Серия КАК Серия,
		|	Остатки.НаСкладе КАК НаСкладе,
		|	Остатки.НаРуках КАК НаРуках,
		|	ЕСТЬNULL(РезервыНаРуках.ДокументРезерв, ЗНАЧЕНИЕ(Документ.торо_РезервПодВнутреннийЗаказ.ПустаяСсылка)) КАК ДокументРезерв,
		|	ЕСТЬNULL(РезервыНаРуках.КоличествоОстаток, 0) КАК РезервНаРуках
		|ИЗ
		|	Остатки КАК Остатки
		|		ЛЕВОЕ СОЕДИНЕНИЕ РезервыНаРуках КАК РезервыНаРуках
		|		ПО Остатки.Склад = РезервыНаРуках.Склад
		|			И Остатки.Номенклатура = РезервыНаРуках.Номенклатура
		|			И Остатки.Характеристика = РезервыНаРуках.Характеристика
		|			И Остатки.Серия = РезервыНаРуках.Серия
		|			И Остатки.ЗаказНаВнутреннееПотребление = РезервыНаРуках.ЗаказНаВнутреннееПотребление";
		РезЗапроса = Запрос.ВыполнитьПакет();
		Выборка = РезЗапроса[2].Выбрать();
		Если НЕ РезЗапроса[3].Пустой() Тогда 
			РезервыНаСкладах = РезЗапроса[3].Выгрузить();
			РезервыНаСкладах.Индексы.Добавить("Склад");
			РезервыНаСкладах.Индексы.Добавить("Номенклатура");
			РезервыНаСкладах.Индексы.Добавить("Характеристика");
			РезервыНаСкладах.Индексы.Добавить("Серия");
		Иначе
			РезервыНаСкладах = Новый ТаблицаЗначений;
		КонецЕсли;
		Если НЕ РезЗапроса[4].Пустой() Тогда 
			РезервыНаРуках = РезЗапроса[4].Выгрузить();
			РезервыНаРуках.Индексы.Добавить("Склад");
			РезервыНаРуках.Индексы.Добавить("Номенклатура");
			РезервыНаРуках.Индексы.Добавить("Характеристика");
			РезервыНаРуках.Индексы.Добавить("Серия");
		Иначе
			РезервыНаРуках = Новый ТаблицаЗначений;
		КонецЕсли;
	Иначе
		Запрос.Текст = "ВЫБРАТЬ
		|	Остатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
		|	Остатки.Склад КАК Склад,
		|	Остатки.Номенклатура КАК Номенклатура,
		|	Остатки.Характеристика КАК Характеристика,
		|	Остатки.Количество КАК Количество,
		|	Остатки.Серия КАК Серия,
		|	Остатки.НаСкладе КАК НаСкладе,
		|	0 КАК НаРуках
		|ИЗ
		|	Остатки КАК Остатки";
		РезЗапроса = Запрос.Выполнить();
		Выборка = РезЗапроса.Выбрать();
		РезервыНаСкладах = Новый ТаблицаЗначений;
		РезервыНаРуках = Новый ТаблицаЗначений;
	КонецЕсли;

	СтруктураПоиска = Новый Структура("Склад, Номенклатура, Характеристика, Серия");  
	Пока Выборка.Следующий() Цикл 
		НадоСписать = Выборка.Количество;
		ЗаполнитьЗначенияСвойств(СтруктураПоиска, Выборка);
		Если Выборка.НаРуках > 0 Тогда
			СписатьОстаток = Мин(НадоСписать, Выборка.НаРуках);
			Движение = Движения.торо_ТоварыНаРуках.ДобавитьРасход();
			Движение.Период = Дата;
			ЗаполнитьЗначенияСвойств(Движение, Выборка);
			Движение.Количество = СписатьОстаток;
			Если РезервыНаРуках.Количество() Тогда
				НайденныеРезервы = РезервыНаРуках.НайтиСтроки(СтруктураПоиска);
				СписатьРезерв = СписатьОстаток;
				Для Каждого СтрокаРезерв Из НайденныеРезервы Цикл
					Если СписатьРезерв = 0 ИЛИ СтрокаРезерв.РезервНаРуках <= 0 Тогда
						Прервать;
					КонецЕсли;
					Движение = Движения.торо_РезервыНаРуках.ДобавитьРасход();
					ЗаполнитьЗначенияСвойств(Движение, СтрокаРезерв);
					Движение.Период = Дата;
					Списать = Мин(СтрокаРезерв.РезервНаРуках, СписатьРезерв);
					Движение.Количество = Списать;
					СписатьРезерв = СписатьРезерв - Списать;
				КонецЦикла;
			КонецЕсли;
			НадоСписать = НадоСписать - СписатьОстаток;
		КонецЕсли;

		Если НадоСписать > 0 Тогда
			Движение = Движения.ТоварыНаСкладах.ДобавитьРасход();
			Движение.Период = Дата;
			ЗаполнитьЗначенияСвойств(Движение, Выборка);
			Движение.ВНаличии = НадоСписать;
			Если РезервыНаСкладах.Количество() Тогда
				НайденныеРезервы = РезервыНаСкладах.НайтиСтроки(СтруктураПоиска);
				СписатьРезерв = НадоСписать;
				Для Каждого СтрокаРезерв Из НайденныеРезервы Цикл
					Если СтрокаРезерв.РезервНаСкладе <= 0 Тогда
						Прервать;
					КонецЕсли;
					Движение = Движения.торо_РезервыНаСкладах.ДобавитьРасход();
					ЗаполнитьЗначенияСвойств(Движение, СтрокаРезерв);
					Движение.Период = Дата;
					Списать = Мин(СписатьРезерв, СтрокаРезерв.РезервНаСкладе);
					Движение.Количество = Списать;
					СписатьРезерв = СписатьРезерв - Списать;
				КонецЦикла;
			КонецЕсли;
		КонецЕсли;
	КонецЦикла;

	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	Остатки.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	Остатки.Склад КАК Склад,
	|	Остатки.Номенклатура КАК Номенклатура,
	|	Остатки.Характеристика КАК Характеристика,
	|	Остатки.Количество КАК Количество,
	|	Остатки.Серия КАК Серия,
	|	Остатки.Назначение КАК Назначение
	|ИЗ
	|	проф_Остатки КАК Остатки";
	РезЗапроса = Запрос.Выполнить();
	Выборка = РезЗапроса.Выбрать();
	РезервыНаСкладах = Новый ТаблицаЗначений;
	РезервыНаРуках = Новый ТаблицаЗначений;	
	
	Пока Выборка.Следующий() Цикл 
		НадоСписать = Выборка.Количество;
		Если НадоСписать > 0 Тогда
			Движение = Движения.проф_ТоварыНаСкладах.ДобавитьРасход();
			Движение.Период = Дата;
			ЗаполнитьЗначенияСвойств(Движение, Выборка);
			Движение.ВНаличии = НадоСписать;
		КонецЕсли;
	КонецЦикла;
	//-- Проф-ИТ, #27, Соловьев А.А., 24.08.2023
	//++ Проф-ИТ, #329, Соловьев А.А., 08.11.2023
	Если Не Отказ Тогда 
		проф_ДвиженияПоРегистру_проф_ОстатокТМЦКСписаниюПоЗакрытымАктамНаРемонт();
	КонецЕсли;
	//-- Проф-ИТ, #329, Соловьев А.А., 08.11.2023
	#КонецВставки
	Движения.Записать();

КонецПроцедуры

//++ Проф-ИТ, #303, Корнилов М.С., 13.10.2023
&После("ПередЗаписью")
Процедура проф_ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	
	Если ПотреблениеПоЗаказам = Ложь Тогда 
		Сообщение = Новый СообщениеПользователю();
		Сообщение.Текст = "Документ должен быть привязан к Заказу на внутреннее потребление!";
		Сообщение.Сообщить();
		Отказ = Истина;
	КонецЕсли;
	
	//++ Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	Если РежимЗаписи = РежимЗаписиДокумента.Проведение Тогда 
		проф_ОбщегоНазначенияВызовСервера.ПроверитьПризнакПодразделенияОрганизации(ЭтотОбъект["Подразделение"], Отказ);
	КонецЕсли;
	//-- Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	
КонецПроцедуры
//-- Проф-ИТ, #303, Корнилов М.С., 13.10.2023

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&ИзменениеИКонтроль("УстановитьБлокировку")
Процедура проф_УстановитьБлокировку()
	Блокировка = Новый БлокировкаДанных;

	ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.ТоварыНаСкладах");
	ЭлементБлокировки.ИсточникДанных = Товары;
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура");
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
	ЭлементБлокировки.УстановитьЗначение("Склад", Склад);
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");

	#Вставка
	//++ Проф-ИТ, #27, Соловьев А.А., 23.08.2023
	ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.проф_ТоварыНаСкладах");
	ЭлементБлокировки.ИсточникДанных = Товары;
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура");
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Назначение", "проф_Назначение");
	ЭлементБлокировки.УстановитьЗначение("Склад", Склад);
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");
	//-- Проф-ИТ, #27, Соловьев А.А., 23.08.2023
	#КонецВставки
	ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.торо_ТоварыНаРуках");
	ЭлементБлокировки.ИсточникДанных = Товары;
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура");
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
	ЭлементБлокировки.УстановитьЗначение("Склад", Склад); 
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");	
	ЭлементБлокировки.УстановитьЗначение("ЗаказНаВнутреннееПотребление", ЗаказНаВнутреннееПотребление);

	ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.торо_РезервыНаСкладах");
	ЭлементБлокировки.ИсточникДанных = Товары;
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура"); 
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
	ЭлементБлокировки.УстановитьЗначение("Склад", Склад);
	ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");

	Если ПотреблениеПоЗаказам Тогда 
		ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.ЗаказыНаВнутреннееПотребление");
		ЭлементБлокировки.ИсточникДанных = Товары;
		ЭлементБлокировки.УстановитьЗначение("ЗаказНаВнутреннееПотребление", ЗаказНаВнутреннееПотребление);
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура");
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
		ЭлементБлокировки.УстановитьЗначение("Склад", Склад);
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");

		ЭлементБлокировки = Блокировка.Добавить("РегистрНакопления.торо_РезервыНаРуках");
		ЭлементБлокировки.ИсточникДанных = Товары;
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Номенклатура", "Номенклатура");
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Характеристика", "Характеристика");
		ЭлементБлокировки.УстановитьЗначение("Склад", Склад);
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Серия", "Серия");	
		ЭлементБлокировки.УстановитьЗначение("ЗаказНаВнутреннееПотребление", ЗаказНаВнутреннееПотребление);
	КонецЕсли;

	Блокировка.Заблокировать();
КонецПроцедуры

&ИзменениеИКонтроль("ЗаполнитьПоЗаказу")
Процедура проф_ЗаполнитьПоЗаказу(ДанныеЗаполнения, РеквизитыШапки, ЗаполнятьНаДату, ОтборТовары)

	ТипДанныеЗаполнения = ТипЗнч(ДанныеЗаполнения);

	Если ТипДанныеЗаполнения = Тип("Массив") Тогда

		ЗаполнитьЗначенияСвойств(ЭтотОбъект, РеквизитыШапки);

	ИначеЕсли ТипДанныеЗаполнения = Тип("ДокументСсылка.ЗаказНаВнутреннееПотребление") Тогда

		ЗаказНаВнутреннееПотребление = ДанныеЗаполнения;

		Запрос = Новый Запрос(
		"ВЫБРАТЬ
		|	Заказы.Ссылка КАК ЗаказНаВнутреннееПотребление,
		|	Заказы.Организация КАК Организация,
		|	Заказы.Подразделение КАК Подразделение,
		|	Заказы.Склад КАК Склад,
		|	Заказы.Склад.УчетныйВидЦены КАК ВидЦены,
		#Вставка
		//++ Проф-ИТ, #183, Соловьев А.А., 15.09.2023
		|	Заказы.проф_НаправлениеДеятельности КАК проф_НаправлениеДеятельности,
		|	Заказы.проф_Назначение КАК проф_Назначение,
		|	Заказы.проф_ХозяйственнаяОперация КАК проф_ХозяйственнаяОперация,
		//-- Проф-ИТ, #183, Соловьев А.А., 15.09.2023
		//++ Проф-ИТ, #27, Соловьев А.А., 25.09.2023
		|	Заказы.ДокументОснование КАК проф_ДокументОснование,
		//-- Проф-ИТ, #27, Соловьев А.А., 25.09.2023
		#КонецВставки
		|	(НЕ Заказы.Проведен) КАК ЕстьОшибкиПроведен
		|ИЗ
		|	Документ.ЗаказНаВнутреннееПотребление КАК Заказы
		|ГДЕ
		|	Заказы.Ссылка = &Заказ");

		Запрос.УстановитьПараметр("Заказ", ЗаказНаВнутреннееПотребление);
		РеквизитыЗаказа = Запрос.Выполнить().Выбрать();
		РеквизитыЗаказа.Следующий();

		МассивДопустимыхСтатусов = Новый Массив();

		ЗаполнитьЗначенияСвойств(ЭтотОбъект, РеквизитыЗаказа);

	КонецЕсли;

	ПотреблениеПоЗаказам = Истина;

	Если ТипДанныеЗаполнения = Тип("Массив") Тогда
		МассивЗаказов = ДанныеЗаполнения;
	Иначе
		МассивЗаказов = Новый Массив;
		МассивЗаказов.Добавить(ЗаказНаВнутреннееПотребление);
	КонецЕсли;

	Документы.ВнутреннееПотреблениеТоваров.ЗаполнитьПоОстаткамЗаказов(ЭтотОбъект,Товары,МассивЗаказов);

	Если ОтборТовары <> Неопределено Тогда
		СтрокиКУдалению = Новый Массив;
		СтруктураПоиска = Новый Структура("ЗаказНаВнутреннееПотребление,КодСтроки");
		Для каждого ДанныеСтроки Из Товары Цикл
			ЗаполнитьЗначенияСвойств(СтруктураПоиска, ДанныеСтроки);
			СписокСтрок = ОтборТовары.НайтиСтроки(СтруктураПоиска);
			Если СписокСтрок.Количество() = 0 Тогда
				СтрокиКУдалению.Добавить(ДанныеСтроки);	
			КонецЕсли; 
		КонецЦикла;
		Для каждого ДанныеСтроки Из СтрокиКУдалению Цикл
			Товары.Удалить(ДанныеСтроки);
		КонецЦикла; 
	КонецЕсли; 

	ЗаполнитьЗаказВШапкеПоЗаказамВТабличнойЧасти(ЗаказНаВнутреннееПотребление, Товары, "ЗаказНаВнутреннееПотребление");

КонецПроцедуры

//++ Проф-ИТ, #303, Корнилов М.С., 13.10.2023
Процедура проф_ПроверитьТЧТоварыИнвентарныйНомер(Отказ)
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	Товары.НомерСтроки КАК НомерСтроки,
		|	Товары.проф_КатегорииЭксплуатации КАК КатегорииЭксплуатации,
		|	Товары.проф_ИнвентарныйНомер КАК ИнвентарныйНомер
		|ПОМЕСТИТЬ ВТ_Товары
		|ИЗ
		|	&Товары КАК Товары
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_Товары.НомерСтроки КАК НомерСтроки
		|ИЗ
		|	ВТ_Товары КАК ВТ_Товары
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.проф_КатегорииЭксплуатации КАК проф_КатегорииЭксплуатации
		|		ПО ВТ_Товары.КатегорииЭксплуатации = проф_КатегорииЭксплуатации.Ссылка
		|			И (проф_КатегорииЭксплуатации.ИнвентарныйУчет = ИСТИНА)
		|			И (ВТ_Товары.ИнвентарныйНомер = """")";
	
	Запрос.УстановитьПараметр("Товары", Товары);	
	ВыборкаТовары = Запрос.Выполнить().Выбрать();
	
	Пока ВыборкаТовары.Следующий() Цикл
		Сообщение = Новый СообщениеПользователю();
		Сообщение.Текст = СтрШаблон("Не заполнена колонка ""Инвентарный номер (Проф-ИТ)"" в строке %1 списка ""Товары""",
			ВыборкаТовары.НомерСтроки);
		Сообщение.Сообщить();
		Отказ = Истина;
	КонецЦикла;
	
КонецПроцедуры	
//-- Проф-ИТ, #303, Корнилов М.С., 13.10.2023

//++ Проф-ИТ, #329, Соловьев А.А., 08.11.2023
Процедура проф_ДвиженияПоРегистру_проф_ОстатокТМЦКСписаниюПоЗакрытымАктамНаРемонт()
	
	Движения.проф_ОстатокТМЦКСписаниюПоЗакрытымАктамНаРемонт.Записывать = Истина;
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	ТЧТовары.Номенклатура КАК Номенклатура,
	|	ТЧТовары.Характеристика КАК Характеристика,
	|	ТЧТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	СУММА(ТЧТовары.Количество) КАК Количество
	|ПОМЕСТИТЬ втТовары
	|ИЗ
	|	Документ.ВнутреннееПотреблениеТоваров.Товары КАК ТЧТовары
	|ГДЕ
	|	ТЧТовары.Ссылка = &Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	ТЧТовары.Номенклатура,
	|	ТЧТовары.Характеристика,
	|	ТЧТовары.ЗаказНаВнутреннееПотребление
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗЛИЧНЫЕ
	|	торо_ИнтеграцияДокументов.ДокументЕРП КАК Заказ,
	|	торо_ИнтеграцияДокументов.ID КАК ID
	|ПОМЕСТИТЬ Заказы
	|ИЗ
	|	РегистрСведений.торо_ИнтеграцияДокументов КАК торо_ИнтеграцияДокументов
	|ГДЕ
	|	торо_ИнтеграцияДокументов.ДокументЕРП В 
	|			(ВЫБРАТЬ 
	|				втТовары.ЗаказНаВнутреннееПотребление 
	|			ИЗ 
	|				втТовары КАК втТовары)
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	торо_ИнтеграцияДокументов.ДокументЕРП КАК ВнутреннееПотребление,
	|	торо_ИнтеграцияДокументов.ID КАК ID
	|ПОМЕСТИТЬ втСоответствиеДокументов
	|ИЗ
	|	РегистрСведений.торо_ИнтеграцияДокументов КАК торо_ИнтеграцияДокументов
	|ГДЕ
	|	торо_ИнтеграцияДокументов.ДокументЕРП = &Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	торо_ИнтеграцияДокументов.ДокументЕРП,
	|	торо_ИнтеграцияДокументов.ID
	|
	|ОБЪЕДИНИТЬ
	|
	|ВЫБРАТЬ РАЗЛИЧНЫЕ
	|	ВнутреннееПотреблениеТоваров.Ссылка,
	|	Заказы.ID
	|ИЗ
	|	Заказы КАК Заказы
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ВнутреннееПотреблениеТоваров.Товары КАК ВнутреннееПотреблениеТоваров
	|		ПО Заказы.Заказ = ВнутреннееПотреблениеТоваров.ЗаказНаВнутреннееПотребление
	|		И  (ВнутреннееПотреблениеТоваров.ЗаказНаВнутреннееПотребление = &Ссылка)
	|ГДЕ
	|	НЕ ВнутреннееПотреблениеТоваров.Ссылка.ПометкаУдаления
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка КАК АктОВыполненииЭтапаРабот,
	|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID КАК ID
	|ПОМЕСТИТЬ втАкты
	|ИЗ
	|	Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
	|ГДЕ
	|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID В
	|			(ВЫБРАТЬ
	|				втСоответствиеДокументов.ID
	|			ИЗ
	|				втСоответствиеДокументов КАК втСоответствиеДокументов)
	|
	|СГРУППИРОВАТЬ ПО
	|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка,
	|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ЗНАЧЕНИЕ(ВидДвиженияНакопления.Расход) КАК ВидДвижения,
	|	втАкты.АктОВыполненииЭтапаРабот КАК АктОВыполненииЭтапаРабот,
	|	втТовары.Номенклатура КАК Номенклатура,
	|	втТовары.Характеристика КАК Характеристика,
	|	СУММА(втТовары.Количество) КАК Количество
	|ИЗ
	|	втТовары КАК втТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ втСоответствиеДокументов КАК втСоответствиеДокументов
	|			ВНУТРЕННЕЕ СОЕДИНЕНИЕ втАкты КАК втАкты
	|			ПО втСоответствиеДокументов.ID = втАкты.ID
	|		ПО (ИСТИНА)
	|
	|СГРУППИРОВАТЬ ПО
	|	втАкты.АктОВыполненииЭтапаРабот,
	|	втТовары.Номенклатура,
	|	втТовары.Характеристика";
	
	Запрос.УстановитьПараметр("Ссылка", Ссылка);
	РезультатЗапроса = Запрос.Выполнить();
	Выборка = РезультатЗапроса.Выбрать();
	
	Пока Выборка.Следующий() Цикл
		
		Движение = Движения.проф_ОстатокТМЦКСписаниюПоЗакрытымАктамНаРемонт.Добавить();
		ЗаполнитьЗначенияСвойств(Движение, Выборка);
		Движение.Период = Дата;
		
	КонецЦикла;
	
КонецПроцедуры
//-- Проф-ИТ, #329, Соловьев А.А., 08.11.2023

#КонецОбласти

