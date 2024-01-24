#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныеПроцедурыИФункции

// Заполняет список команд печати.
//
// Параметры:
// КомандыПечати – ТаблицаЗначений – состав полей см. в функции УправлениеПечатью.СоздатьКоллекциюКомандПечати.
//
Процедура ДобавитьКомандыПечати(КомандыПечати) Экспорт
	
	КомандаПечати = КомандыПечати.Добавить();
	КомандаПечати.МенеджерПечати = "Документ.торо_ПринятиеОРКУчету";
	КомандаПечати.Идентификатор = "ПринятиеОРКУчету";
	КомандаПечати.Представление = НСтр("ru = 'Принятие объекта ремонта к учету'");
	КомандаПечати.СразуНаПринтер = ОбщегоНазначенияВызовСервера.ХранилищеОбщихНастроекЗагрузить(
	"НастройкиТОиР",
	"ПечатьДокументовБезПредварительногоПросмотра",
	Ложь);
	
КонецПроцедуры

// Сформировать печатные формы объектов.
//
// ВХОДЯЩИЕ:
//   ИменаМакетов    - Строка    - Имена макетов, перечисленные через запятую.
//   МассивОбъектов  - Массив    - Массив ссылок на объекты которые нужно распечатать.
//   ПараметрыПечати - Структура - Структура дополнительных параметров печати.
//
// ИСХОДЯЩИЕ:
//   КоллекцияПечатныхФорм - Таблица значений - Сформированные табличные документы.
//   ПараметрыВывода       - Структура        - Параметры сформированных табличных документов.
//
Процедура Печать(МассивОбъектов, ПараметрыПечати, КоллекцияПечатныхФорм, ОбъектыПечати, ПараметрыВывода) Экспорт
	
	Если УправлениеПечатью.НужноПечататьМакет(КоллекцияПечатныхФорм, "ПринятиеОРКУчету") Тогда
		
		УправлениеПечатью.ВывестиТабличныйДокументВКоллекцию(КоллекцияПечатныхФорм, 
		"ПринятиеОРКУчету", 
		"Принятие объекта ремонта к учету", 
		ПечатьДокументаПринятиеОРКУчету(МассивОбъектов, ПараметрыПечати),
		,
		"Документ.торо_ПринятиеОРКУчету.ПФ_MXL_ПринятиеОРКУчету");
		
	КонецЕсли;
	
КонецПроцедуры

Функция ПечатьДокументаПринятиеОРКУчету(МассивОбъектов, ПараметрыПечати)
	
	ТабДокумент = Новый ТабличныйДокумент;
	Макет = УправлениеПечатью.МакетПечатнойФормы("Документ.торо_ПринятиеОРКУчету.ПФ_MXL_ПринятиеОРКУчету");
	
	Запрос = Новый Запрос;
	Запрос.Текст =  "ВЫБРАТЬ
	                |	торо_ПринятиеОРКУчету.Ссылка КАК Документ,
	                |	торо_ПринятиеОРКУчету.Номер КАК Номер,
	                |	торо_ПринятиеОРКУчету.Дата КАК Дата,
	                |	торо_ПринятиеОРКУчету.ОбъектРемонта КАК ОбъектРемонта,
	                |	торо_ПринятиеОРКУчету.ОбъектРемонта.ИнвентарныйНомер КАК ИнвентарныйНомер,
	                |	торо_ПринятиеОРКУчету.Организация КАК Организация,
	                |	торо_ПринятиеОРКУчету.Подразделение КАК Подразделение,
	                |	торо_ПринятиеОРКУчету.ДатаВводаВЭксплуатацию КАК ДатаВводаВЭксплуатацию,
	                |	торо_ПринятиеОРКУчету.ВходитВСостав КАК ВходитВСостав,
	                |	торо_ПринятиеОРКУчету.Проведен КАК Проведен
	                |ПОМЕСТИТЬ Данные
	                |ИЗ
	                |	Документ.торо_ПринятиеОРКУчету КАК торо_ПринятиеОРКУчету
	                |ГДЕ
	                |	торо_ПринятиеОРКУчету.Ссылка В(&Ссылка)
	                |
	                |ОБЪЕДИНИТЬ ВСЕ
	                |
	                |ВЫБРАТЬ
	                |	торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету.Ссылка,
	                |	NULL,
	                |	NULL,
	                |	торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету.ОбъектРемонта,
	                |	торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету.ОбъектРемонта.ИнвентарныйНомер,
	                |	NULL,
	                |	NULL,
	                |	NULL,
	                |	NULL,
	                |	NULL
	                |ИЗ
	                |	Документ.торо_ПринятиеОРКУчету.СписокПодчиненныхПринятыхКУчету КАК торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету
	                |ГДЕ
	                |	торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету.Ссылка В(&Ссылка)
	                |	И торо_ПринятиеОРКУчетуСписокПодчиненныхПринятыхКУчету.ПринятьКУчету
	                |;
	                |
	                |////////////////////////////////////////////////////////////////////////////////
	                |ВЫБРАТЬ
	                |	Данные.Документ КАК Документ,
	                |	Данные.Номер КАК Номер,
	                |	Данные.Дата КАК Дата,
	                |	Данные.ОбъектРемонта КАК ОбъектРемонта,
	                |	Данные.ИнвентарныйНомер КАК ИнвентарныйНомер,
	                |	Данные.Организация КАК Организация,
	                |	Данные.Подразделение КАК Подразделение,
	                |	Данные.ДатаВводаВЭксплуатацию КАК ДатаВводаВЭксплуатацию,
	                |	Данные.ВходитВСостав КАК ВходитВСостав,
	                |	Данные.Проведен КАК Проведен,
	                |	Данные.ДатаВводаВЭксплуатацию ЕСТЬ NULL КАК ЭтоПодчиненный
	                |ИЗ
	                |	Данные КАК Данные
	                |
	                |УПОРЯДОЧИТЬ ПО
	                |	ЭтоПодчиненный
	                |ИТОГИ
	                |	МАКСИМУМ(Номер),
	                |	МАКСИМУМ(Дата),
	                |	МАКСИМУМ(Организация),
	                |	МАКСИМУМ(Подразделение)
	                |ПО
	                |	Документ";
	
	Запрос.УстановитьПараметр("Ссылка", МассивОбъектов);
	РезЗапроса = Запрос.Выполнить();
	ВыборкаДокументов = РезЗапроса.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
	
	Пока ВыборкаДокументов.Следующий() Цикл
		
		ОбластьМакета = Макет.ПолучитьОбласть("Шапка");
		ОбластьМакета.Параметры.Организация = торо_ЗаполнениеДокументов.ПолучитьПредставлениеОрганизацииДляПечати(ВыборкаДокументов.Организация);
		ОбластьМакета.Параметры.Подразделение = ВыборкаДокументов.Подразделение;
		ОбластьМакета.Параметры.Номер = ВыборкаДокументов.Номер;
		ОбластьМакета.Параметры.Дата = ВыборкаДокументов.Дата;
		ТабДокумент.Вывести(ОбластьМакета);
		
		ВыборкаОР = ВыборкаДокументов.Выбрать();
		КоличествоПодчиненных = 0;
		Пока ВыборкаОР.Следующий() Цикл 
			Если ВыборкаОР.ЭтоПодчиненный Тогда
				Если КоличествоПодчиненных = 0 Тогда
					ОбластьШапкаПодчиненных = Макет.ПолучитьОбласть("ПодчиненныеШапка");
					ТабДокумент.Вывести(ОбластьШапкаПодчиненных);
				КонецЕсли;
				КоличествоПодчиненных = КоличествоПодчиненных + 1;
				ОбластьМакета = Макет.ПолучитьОбласть("ПодчиненныйОРСтрока");
				ОбластьМакета.Параметры.НомерСтроки = КоличествоПодчиненных;
			Иначе
				ОбластьМакета = Макет.ПолучитьОбласть("ОбщиеСведения");
				ОбластьМакета.Параметры.ДатаВвода = ВыборкаОР.ДатаВводаВЭксплуатацию;
				ОбластьМакета.Параметры.Иерархия = ПреобразоватьСтрокуИерархии(ВыборкаОР.ВходитВСостав);
			КонецЕсли;
			ОбластьМакета.Параметры.ОбъектРемонта = ВыборкаОР.ОбъектРемонта;
			ОбластьМакета.Параметры.ИнвентарныйНомер = ВыборкаОР.ИнвентарныйНомер;
			ТабДокумент.Вывести(ОбластьМакета);
		КонецЦикла;
		
		Если КоличествоПодчиненных Тогда 
			ОбластьМакета = Макет.ПолучитьОбласть("ПодчиненныеПодвал");
	        ТабДокумент.Вывести(ОбластьМакета);
		КонецЕсли;
		
		ВывестиКомиссию(ВыборкаДокументов.Документ, Макет, ТабДокумент);
		
		ТабДокумент.ВывестиГоризонтальныйРазделительСтраниц();
	КонецЦикла;
	
	ТабДокумент.ТолькоПросмотр = Истина;
	ТабДокумент.КлючПараметровПечати = "торо_ПечатьДокументаПринятиеОРКУчету";
	Возврат ТабДокумент;
	
КонецФункции

Функция ПреобразоватьСтрокуИерархии(ИсходныеСтроки)
	
	МассивСтрок = СтрРазделить(ИсходныеСтроки, Символы.ПС);
	Индекс = 0;
	СтрокаИерархии= "";
	Для каждого Стр Из МассивСтрок Цикл
		СтрокаИерархии = СтрокаИерархии + Сред(Стр, Индекс*4 + 1) + " - ";
		Индекс = Индекс + 1;
	КонецЦикла;
	СтрокаИерархии = Лев(СтрокаИерархии, СтрДлина(СтрокаИерархии) - 3);
	
	Возврат СтрокаИерархии;
	
КонецФункции

Процедура ВывестиКомиссию(СсылкаНаДокумент, Макет, ТабДокумент)
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	                |	торо_ПринятиеОРКУчетуСотрудники.Сотрудник КАК Сотрудник
	                |ИЗ
	                |	Документ.торо_ПринятиеОРКУчету.Сотрудники КАК торо_ПринятиеОРКУчетуСотрудники
	                |ГДЕ
	                |	торо_ПринятиеОРКУчетуСотрудники.Ссылка = &Ссылка";
	
	Запрос.УстановитьПараметр("Ссылка", СсылкаНаДокумент);
	
	Результат = Запрос.Выполнить();
	ВыборкаКомиссия = Результат.Выбрать();
	
	Если Результат.Пустой() Тогда
		Возврат;
	КонецЕсли;
	
	ОбластьМакета = Макет.ПолучитьОбласть("КомиссияШапка");
    ТабДокумент.Вывести(ОбластьМакета);
	
	ОбластьМакета = Макет.ПолучитьОбласть("Комиссия");
	Пока ВыборкаКомиссия.Следующий() Цикл
		 ОбластьМакета.Параметры.Сотрудник = ВыборкаКомиссия.Сотрудник;
		 ТабДокумент.Вывести(ОбластьМакета);		
	КонецЦикла;
		
КонецПроцедуры

Процедура ПриОпределенииНастроекВерсионированияОбъектов(Настройки) Экспорт
	
КонецПроцедуры

// СтандартныеПодсистемы.УправлениеДоступом

// См. УправлениеДоступомПереопределяемый.ПриЗаполненииСписковСОграничениемДоступа.
Процедура ПриЗаполненииОграниченияДоступа(Ограничение) Экспорт
	
	Ограничение.Текст =
	"РазрешитьЧтение
	|ГДЕ 
	|	ЗначениеРазрешено(Организация)
	|	И ЗначениеРазрешено(Подразделение)
	|;
	|РазрешитьИзменениеЕслиРазрешеноЧтение
	|ГДЕ
	|	(&Ограничение_ОР)";

	ОграничениеОР = торо_УправлениеДоступом.ПолучитьОграничениеДоступаДляОбъектаРемонтаВТаблице("ОбъектРемонта", Истина);
	Ограничение.Текст = СтрЗаменить(Ограничение.Текст, "&Ограничение_ОР", ОграничениеОР);

	Ограничение.ТекстДляВнешнихПользователей = Ограничение.Текст;

КонецПроцедуры

// Конец СтандартныеПодсистемы.УправлениеДоступом

#КонецОбласти

#КонецЕсли
