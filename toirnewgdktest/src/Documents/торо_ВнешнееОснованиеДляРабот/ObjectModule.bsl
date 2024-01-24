#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

Перем СтруктураДанных Экспорт; // Содержит структуру данных о статусе документа (Начальный, СогласованиеЗавершено, ПроводитьДокумент, ШаблонСообщенияЭлектроннойПочты) 
										 // из регистра сведений торо_МатрицаПереходаСтатусовДокументов. 
Перем БезусловнаяЗапись Экспорт; // Отключает проверки при записи документа

#Область ОбработчикиСобытий

Процедура ПриКопировании(ОбъектКопирования)
	
	Для каждого ТекСтрока из ОбследованноеОборудование Цикл
		
		МассивОпераций = РемонтныеРаботы.НайтиСтроки(Новый Структура("Предписание_ID", ТекСтрока.ID));
		
		ТекСтрока.ID = Новый УникальныйИдентификатор();
		
		Для каждого Операция Из МассивОпераций Цикл
			Операция.Предписание_ID = ТекСтрока.ID;
		КонецЦикла;
		
	КонецЦикла;
	
КонецПроцедуры

Процедура ОбработкаЗаполнения(Основание, СтандартнаяОбработка)
	
	Если ТипЗнч(Основание) = Тип("Структура") Тогда
		ЗаполнитьЗначенияСвойств(ЭтотОбъект, Основание);
		
		Если Основание.Свойство("МассивСтрокТЧОбследованноеОборудование") Тогда
			Для каждого ТекущаяСтрока Из Основание.МассивСтрокТЧОбследованноеОборудование Цикл
				НоваяСтрока = ОбследованноеОборудование.Добавить();
				ЗаполнитьЗначенияСвойств(НоваяСтрока, ТекущаяСтрока);
			КонецЦикла;
		КонецЕсли;
		
		Если Основание.Свойство("МассивСтрокТЧРемонтныеРаботы") Тогда
			Для каждого ТекущаяСтрока Из Основание.МассивСтрокТЧРемонтныеРаботы Цикл
				НоваяСтрока = РемонтныеРаботы.Добавить();
				ЗаполнитьЗначенияСвойств(НоваяСтрока, ТекущаяСтрока);
			КонецЦикла;
		КонецЕсли;
	КонецЕсли;
	
	Если ТипЗнч(Основание) = Тип("ДокументСсылка.торо_АктОВыполненииРегламентногоМероприятия") Тогда
		Организация   = Основание.Организация;
		Подразделение = Основание.Подразделение;
		ДокументОснование = Основание;		
	КонецЕсли;
	
	Если ТипЗнч(Основание) = Тип("ДокументСсылка.торо_АнализКоренныхПричин") Тогда
		Организация   = Основание.Организация;
		Подразделение = Основание.Подразделение;
		торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
		
		Неотмененный = торо_ЗаполнениеДокументов20.ВернутьIDНеобработанногоРемонта(Основание);
		Если Неотмененный <> "" Тогда
			ЗаполнитьВнешнееОснованиеПоАнализу(Основание, ЭтотОбъект, Неотмененный);	
		КонецЕсли;
		ДокументОснование = Основание;
	КонецЕсли;	
	
	Если НЕ ЗначениеЗаполнено(ДатаСоздания) И ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
		"НастройкиТОиР", 
		"ПодставлятьТекущуюДатуВоВнешнееОснованиеИВыявленныеДефекты",
		Истина) Тогда
		
		ДатаСоздания = ТекущаяДата();
		
	КонецЕсли;
	
	торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
	
КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	Если ОбследованноеОборудование.Количество() = 0 Тогда
		ТекстСообщения = НСтр("ru = 'В документе не заполнена табличная часть обследованного оборудования.'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,, Отказ);
		Возврат;
	КонецЕсли;
	
	ПроверитьЗаполнениеТабличнойЧастиОбследованноеОборудование(Отказ);
	ПроверитьЗаполнениеТабличнойЧастиРемонтныеРаботы(Отказ);
	
	ПроверитьЗаполнениеПриВводеНаОснованииАнализа(Отказ);
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, Режим)
	
	ПроверитьЗаполнениеПриВводеНаОснованииАнализа(Отказ);
	
	// регистр торо_Предписания
	Движения.торо_ВнешниеОснованияДляРабот.Записывать = Истина;
	Движения.торо_ВнешниеОснованияДляРабот.Очистить();
	
	Для Каждого ТекСтрокаОбследованноеОборудование Из ОбследованноеОборудование Цикл
		// регистр торо_Предписания
		Движение = Движения.торо_ВнешниеОснованияДляРабот.Добавить();
		ЗаполнитьЗначенияСвойств(Движение, ТекСтрокаОбследованноеОборудование);
		Движение.Период                    = Дата;
		Движение.Организация               = Организация;
		Движение.Подразделение             = Подразделение;
	КонецЦикла;

	// регистр торо_ОперацииПредписаний
	Движения.торо_ОперацииВнешнихОснованийДляРабот.Записывать = Истина;
	Движения.торо_ОперацииВнешнихОснованийДляРабот.Очистить();
	
	Для Каждого ТекСтрокаРемонтныеРаботы Из РемонтныеРаботы Цикл
		Движение = Движения.торо_ОперацииВнешнихОснованийДляРабот.Добавить();
		Движение.Предписание_ID  = ТекСтрокаРемонтныеРаботы.Предписание_ID;
		Движение.РемонтнаяРабота = ТекСтрокаРемонтныеРаботы.РемонтнаяРабота;
		Движение.ID              = ТекСтрокаРемонтныеРаботы.ID;
		Движение.Родитель_ID     = ТекСтрокаРемонтныеРаботы.Родитель_ID;
		Движение.Количество      = ТекСтрокаРемонтныеРаботы.Количество;
		Движение.ID_ПараллельнойОперации = ТекСтрокаРемонтныеРаботы.ID_ПараллельнойОперации;
	КонецЦикла;

	ДвиженияПоРегистру_торо_ГарантийностьРемонтов();
	ДвиженияПоРегистру_торо_ВсеДокументыПоРемонтам();
	
	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусРемонтовПриПроведении(ЭтотОбъект);
	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусДокумента(Ссылка, Ссылка, Перечисления.торо_СтатусыДокументов.Зарегистрирован);
	
	торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.Проведение);
	
КонецПроцедуры

Процедура ОбработкаУдаленияПроведения(Отказ)
	
	МассивIDДляБлокировки = торо_Ремонты.МассивIDДляБлокировкиРемонтовОборудования(Ссылка, ОбследованноеОборудование.ВыгрузитьКолонку("ID"));
	Если МассивIDДляБлокировки <> Неопределено И МассивIDДляБлокировки.Количество() > 0 Тогда
		ТекстСообщения = НСтр("ru = 'Невозможно отменить проведение документа, так как имеются созданные на его основании проведенные документы!'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,, Отказ);
	КонецЕсли;
	
	Если Не Отказ Тогда
		торо_РаботаСоСтатусамиДокументовСервер.ОтменаПроведения(Ссылка);
		торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.ОтменаПроведения);
	КонецЕсли;
	
КонецПроцедуры

Процедура ПриЧтенииПредставленийНаСервере() Экспорт
    МультиязычностьСервер.ПриЧтенииПредставленийНаСервере(ЭтотОбъект);
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	
	МассивОР = ОбщегоНазначения.ВыгрузитьКолонку(ОбследованноеОборудование, "ОбъектРемонта", Истина);
	ОбъектыРемонтаСтрокой = торо_ЗаполнениеДокументов20.СформироватьСтрокуОбъектовРемонта(МассивОР);
	
	// Получение удаленных ремонтов при перепроведении документа. Используются при формировании записей
	// регистров торо_ПлановыеИсполнителиРемонтов, торо_ОбщиеДанныеПоРемонтам, торо_СвернутыеФактическиеДатыРемонтов.
	Если Проведен И РежимЗаписи = РежимЗаписиДокумента.Проведение Тогда
		МассивIDДокумента = ОбследованноеОборудование.ВыгрузитьКолонку("ID");
		МассивУдаленныхID = торо_Ремонты.ПолучитьIDУдаленныхРемонтовДокумента(МассивIDДокумента, Ссылка);
		ДополнительныеСвойства.Вставить("МассивУдаленныхID", МассивУдаленныхID);
	Иначе
		ДополнительныеСвойства.Вставить("МассивУдаленныхID", Новый Массив());
	КонецЕсли;
	
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ДвиженияПоРегистру_торо_ВсеДокументыПоРемонтам()
	Движения.торо_ВсеДокументыПоРемонтам.Записывать = Истина;
	Движения.торо_ВсеДокументыПоРемонтам.Очистить();
	
	Для каждого СтрокаТЧ Из ОбследованноеОборудование Цикл
	    Движение = Движения.торо_ВсеДокументыПоРемонтам.Добавить();
		Движение.IDРемонта = СтрокаТЧ.ID;
		Движение.ДокументПоРемонту = Ссылка;
		Движение.ДатаДокумента = Дата;
		Движение.Организация = Организация;
		Движение.Подразделение = Подразделение;
	КонецЦикла;
КонецПроцедуры

Процедура ДвиженияПоРегистру_торо_ГарантийностьРемонтов()
	
	Если НЕ ПолучитьФункциональнуюОпцию("торо_ИспользоватьГарантии") Тогда
		Возврат;
	КонецЕсли;
	
	Движения.торо_ГарантийностьРемонтов.Записывать = Истина;
	Движения.торо_ГарантийностьРемонтов.Очистить();
	
	Для каждого Строка Из ОбследованноеОборудование Цикл
		Движение = Движения.торо_ГарантийностьРемонтов.Добавить();
		Движение.Период = Дата;
		Движение.IDРемонта = Строка.ID;
		Движение.Гарантийный = Строка.ГарантийныйРемонт;
	КонецЦикла;
	
КонецПроцедуры

Процедура ПроверитьЗаполнениеТабличнойЧастиОбследованноеОборудование(Отказ)
	
	// Укажем, что надо проверить:
	СтруктураОбязательныхПолей = Новый Структура("ОбъектРемонта, ПлановаяДатаРемонта, Ответственный", "", "не заполнено значение реквизита ""Срок""!");
	
	// Вызовем общую процедуру проверки.
	торо_ЗаполнениеДокументов.ПроверитьЗаполнениеТабличнойЧасти(ЭтотОбъект, "ОбследованноеОборудование", СтруктураОбязательныхПолей, Отказ, "");
	
КонецПроцедуры

Процедура ПроверитьЗаполнениеТабличнойЧастиРемонтныеРаботы(Отказ)
	
	Запрос = Новый Запрос("ВЫБРАТЬ
	|	РемонтныеРаботы.Предписание_ID
	|ПОМЕСТИТЬ РемонтныеРаботы
	|ИЗ
	|	&РемонтныеРаботы КАК РемонтныеРаботы
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ОбследованноеОборудование.НомерСтроки,
	|	ОбследованноеОборудование.ID,
	|	ОбследованноеОборудование.ОбъектРемонта,
	|	ОбследованноеОборудование.ПлановаяДатаРемонта
	|ПОМЕСТИТЬ ОбследованноеОборудование
	|ИЗ
	|	&ОбследованноеОборудование КАК ОбследованноеОборудование
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ОбследованноеОборудование.НомерСтроки,
	|	ОбследованноеОборудование.ОбъектРемонта
	|ИЗ
	|	ОбследованноеОборудование КАК ОбследованноеОборудование
	|		ЛЕВОЕ СОЕДИНЕНИЕ РемонтныеРаботы КАК РемонтныеРаботы
	|		ПО ОбследованноеОборудование.ID = РемонтныеРаботы.Предписание_ID
	|ГДЕ
	|	РемонтныеРаботы.Предписание_ID ЕСТЬ NULL ");
	Запрос.УстановитьПараметр("ОбследованноеОборудование", ОбследованноеОборудование);
	Запрос.УстановитьПараметр("РемонтныеРаботы", РемонтныеРаботы);
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		ШаблонСообщения = НСтр("ru = 'Для объекта ремонта %1 в строке %2 не указано ни одной рекомендуемой операции.'");
		ТекстСообщения = СтрШаблон(ШаблонСообщения, Выборка.ОбъектРемонта, Выборка.НомерСтроки); 
		Поле = СтрШаблон("Объект.ОбследованноеОборудование[%1].ОбъектРемонта", Выборка.НомерСтроки - 1); 
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
	КонецЦикла;
	
КонецПроцедуры

Процедура ПроверитьЗаполнениеПриВводеНаОснованииАнализа(Отказ)
	Если ТипЗнч(ДокументОснование) = Тип("ДокументСсылка.торо_АнализКоренныхПричин") Тогда
		Запрос = Новый Запрос;
		Запрос.Текст = 
			"ВЫБРАТЬ
			|	торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.Ссылка КАК Ссылка,
			|	торо_АнализКоренныхПричинКорректирующиеМероприятия.ИДМероприятия КАК ИДМероприятия,
			|	торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.ОбъектРемонта КАК ОбъектРемонта,
			|	торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.Ссылка.ДокументОснование КАК ДокументОснование
			|ИЗ
			|	Документ.торо_ВнешнееОснованиеДляРабот.ОбследованноеОборудование КАК торо_ВнешнееОснованиеДляРаботОбследованноеОборудование
			|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АнализКоренныхПричин.КорректирующиеМероприятия КАК торо_АнализКоренныхПричинКорректирующиеМероприятия
			|		ПО торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.Ссылка.ДокументОснование = торо_АнализКоренныхПричинКорректирующиеМероприятия.Ссылка
			|			И торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.ID = торо_АнализКоренныхПричинКорректирующиеМероприятия.ИДМероприятия
			|ГДЕ
			|	торо_ВнешнееОснованиеДляРаботОбследованноеОборудование.Ссылка = &Ссылка
			|	И торо_АнализКоренныхПричинКорректирующиеМероприятия.ИДМероприятия ЕСТЬ NULL";
		
		Запрос.УстановитьПараметр("Ссылка", Ссылка);
		Результат = Запрос.Выполнить();
		
		ВыборкаДетальныеЗаписи = Результат.Выбрать();
		
		Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
			СтрокаСообщения = СтрШаблон(НСтр("ru = 'Ремонт в документе-основании ""%1"" для объектра ремонта ""%2"" был удален'"), 
				ВыборкаДетальныеЗаписи.ДокументОснование, ВыборкаДетальныеЗаписи.ОбъектРемонта);
			ОбщегоНазначения.СообщитьПользователю(СтрокаСообщения,,,, Отказ);		
		КонецЦикла;
	КонецЕсли; 
КонецПроцедуры

Процедура ЗаполнитьВнешнееОснованиеПоАнализу(Анализ, ВнешнееОснование, Неотмененный)

	ЗаполнитьЗначенияСвойств(ВнешнееОснование, Анализ);
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.Описание КАК Описание,
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.Ответственный КАК Ответственный,
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.СрокВыполнения КАК ПлановаяДатаРемонта,
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.ИДМероприятия КАК ID,
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.Ссылка.ОбъектРемонта КАК ОбъектРемонта
		|ИЗ
		|	Документ.торо_АнализКоренныхПричин.КорректирующиеМероприятия КАК торо_АнализКоренныхПричинКорректирующиеМероприятия
		|ГДЕ
		|	торо_АнализКоренныхПричинКорректирующиеМероприятия.ИДМероприятия = &ИДМероприятия
		|	И торо_АнализКоренныхПричинКорректирующиеМероприятия.Ссылка = &Ссылка";
	
	Запрос.УстановитьПараметр("ИДМероприятия", Неотмененный);
	Запрос.УстановитьПараметр("Ссылка", Анализ);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		НоваяСтрока = ВнешнееОснование.ОбследованноеОборудование.Добавить();
		ЗаполнитьЗначенияСвойств(НоваяСтрока, ВыборкаДетальныеЗаписи);
	КонецЦикла;

КонецПроцедуры

БезусловнаяЗапись = Ложь;

#КонецОбласти

#КонецЕсли