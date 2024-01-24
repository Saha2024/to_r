#Если Сервер Или ТолстыйКлиентОбычноеПриложение Тогда

#Область СлужебныйПрограммныйИнтерфейс

Функция РезультатАктуализацииКешаКонтрагентов() Экспорт
	
	Результат = Новый Структура;
	Результат.Вставить("Успешно",           Истина);
	Результат.Вставить("КоличествоЗаписей", 0);
	Результат.Вставить("ТекстОшибки",       0);
	
	УстановитьПривилегированныйРежим(Истина);
	
	ИмяСправочникаКонтрагенты = ОбщегоНазначенияБЭД.ИмяПрикладногоСправочника("Контрагенты");
	РеквизитИНН = ОбщегоНазначенияБЭДПовтИсп.ИмяНаличиеОбъектаРеквизитаВПрикладномРешении("ИННКонтрагента");
	РеквизитКПП = ОбщегоНазначенияБЭДПовтИсп.ИмяНаличиеОбъектаРеквизитаВПрикладномРешении("КППКонтрагента");
	
	ТаблицаКОбновлению = Новый ТаблицаЗначений;
	ТаблицаКОбновлению.Колонки.Добавить("Контрагент", Новый ОписаниеТипов("СправочникСсылка." + ИмяСправочникаКонтрагенты));
	ТаблицаКОбновлению.Колонки.Добавить("ИНН",           Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(12)));
	ТаблицаКОбновлению.Колонки.Добавить("КПП",           Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(9)));
	ТаблицаКОбновлению.Колонки.Добавить("Идентификатор", Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(36)));
	
	ТаблицаДанныеСервиса = Новый ТаблицаЗначений;
	ТаблицаДанныеСервиса.Колонки.Добавить("ИНН",           Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(12)));
	ТаблицаДанныеСервиса.Колонки.Добавить("КПП",           Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(9)));
	ТаблицаДанныеСервиса.Колонки.Добавить("Идентификатор", Новый ОписаниеТипов("Строка",, Новый КвалификаторыСтроки(36)));

	
	Запрос = Новый Запрос;
	ТекстЗапроса = "ВЫБРАТЬ
	|	Контрагенты.Ссылка                  КАК Контрагент,
	|	Контрагенты.ИНН                     КАК ИНН,
	|	Контрагенты.КПП                     КАК КПП,
	|	КонтрагентыБизнесСеть.Идентификатор КАК Идентификатор
	|ИЗ
	|	&Контрагенты КАК Контрагенты
	|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.КонтрагентыБизнесСеть КАК КонтрагентыБизнесСеть
	|		ПО Контрагенты.Ссылка = КонтрагентыБизнесСеть.Контрагент
	|ГДЕ
	|	Контрагенты.ИНН <> """"
	|	И НЕ Контрагенты.ПометкаУдаления
	|	И КонтрагентыБизнесСеть.Идентификатор IS NULL";
	
	
	ТекстЗапроса = СтрЗаменить(ТекстЗапроса, "&Контрагенты", "Справочник." +     ИмяСправочникаКонтрагенты);
	ТекстЗапроса = СтрЗаменить(ТекстЗапроса, "Контрагенты.ИНН", "Контрагенты." + РеквизитИНН);
	ТекстЗапроса = СтрЗаменить(ТекстЗапроса, "Контрагенты.КПП", "Контрагенты." + РеквизитКПП);

	Запрос.Текст = ТекстЗапроса;
	
	РезультатЗапроса = Запрос.Выполнить();
	Если РезультатЗапроса.Пустой() Тогда
		Возврат Результат;
	КонецЕсли;
	
	Выборка = РезультатЗапроса.Выбрать();
	
	КоличествоКонтрагентов = 0;
	НатуральныеИдентификаторы = Новый Массив;
	
	Пока Выборка.Следующий() Цикл
		
		Отказ = Ложь;
		КортежИдентификаторов = БизнесСеть.ИдентификаторыУчастника(Выборка.ИНН,
		                                                           Выборка.КПП,
		                                                           ,
		                                                           Отказ);
		
		Если Отказ Тогда
			Продолжить;
		КонецЕсли;
		
		ДобавитьЗаписьТаблицуКонтрагентовПриПолученииИдентификаторов(ТаблицаКОбновлению, Выборка);
		Попытка 
			НатуральныйИдентификатор = БизнесСеть.НатуральныйИдентификаторОрганизации(КортежИдентификаторов);
		Исключение
			Продолжить;
		КонецПопытки;
		НатуральныеИдентификаторы.Добавить(НатуральныйИдентификатор);
		
		КоличествоКонтрагентов = КоличествоКонтрагентов + 1;
		
		Если НатуральныеИдентификаторы.Количество() = 200 Тогда
			ДополнитьТаблицуИдентификаторовПоДаннымСервиса(ТаблицаДанныеСервиса, НатуральныеИдентификаторы);
			КоличествоКонтрагентов = 0;
		КонецЕсли;
		
	КонецЦикла;
	
	Если НатуральныеИдентификаторы.Количество() > 0 Тогда
		ДополнитьТаблицуИдентификаторовПоДаннымСервиса(ТаблицаДанныеСервиса, НатуральныеИдентификаторы);
	КонецЕсли;
	
	Если ТаблицаДанныеСервиса.Количество() > 0 Тогда
		ОбработатьПолученныеДанныеИдентификаторовКонтрагентовБизнесСети(ТаблицаКОбновлению, ТаблицаДанныеСервиса, Результат);
	КонецЕсли;
	
	Возврат Результат;
	
КонецФункции

Функция КонтрагентыEDIПоОрганизации() Экспорт
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	КонтрагентыБизнесСеть.Контрагент КАК Контрагент
	|ИЗ
	|	РегистрСведений.КонтрагентыБизнесСеть КАК КонтрагентыБизнесСеть
	|";
	
	Возврат Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Контрагент");
	
КонецФункции

Процедура ОбновитьКэшДанныхКонтрагентаБизнесСети(Знач ДанныеКонтрагентов, Отказ, ТекстОшибки = "") Экспорт
	
	Если ДанныеКонтрагентов.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	
	НачатьТранзакцию();
	Попытка
		Блокировка = Новый БлокировкаДанных;
		ЭлементБлокировки = Блокировка.Добавить("РегистрСведений.КонтрагентыБизнесСеть");
		ЭлементБлокировки.Режим = РежимБлокировкиДанных.Исключительный;
		ЭлементБлокировки.ИсточникДанных = ДанныеКонтрагентов;
		ЭлементБлокировки.ИспользоватьИзИсточникаДанных("Контрагент", "Контрагент");
		Блокировка.Заблокировать();
		
		Для Каждого ДанныеКонтрагента Из ДанныеКонтрагентов Цикл
			Если Не ЗначениеЗаполнено(ДанныеКонтрагента.Контрагент) Тогда
				Продолжить;
			КонецЕсли;
			
			НаборЗаписей = РегистрыСведений.КонтрагентыБизнесСеть.СоздатьНаборЗаписей();
			НаборЗаписей.Отбор.Контрагент.Установить(ДанныеКонтрагента.Контрагент);
			НоваяЗапись = НаборЗаписей.Добавить();
			ЗаполнитьЗначенияСвойств(НоваяЗапись, ДанныеКонтрагента);
			НаборЗаписей.Записать();
			
		КонецЦикла;
		ЗафиксироватьТранзакцию();
	Исключение
		ОтменитьТранзакцию();
		
		Отказ = Истина;
		
		ТекстОшибки = ПодробноеПредставлениеОшибки(ИнформацияОбОшибке());
	КонецПопытки
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ОбработатьПолученныеДанныеИдентификаторовКонтрагентовБизнесСети(ТаблицаКОбновлению, ТаблицаДанныеСервиса, Результат)
	
	Запрос = Новый Запрос;
	
	Запрос.Текст = "ВЫБРАТЬ
	|	ТаблицаКОбновлению.Контрагент    КАК Контрагент,
	|	ТаблицаКОбновлению.ИНН           КАК ИНН,
	|	ТаблицаКОбновлению.КПП           КАК КПП,
	|	ТаблицаКОбновлению.Идентификатор КАК Идентификатор
	|ПОМЕСТИТЬ ТаблицаКОбновлению
	|ИЗ
	|	&ТаблицаКОбновлению КАК ТаблицаКОбновлению
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ИНН,
	|	КПП
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ДанныеСервиса.Идентификатор КАК Идентификатор,
	|	ДанныеСервиса.ИНН           КАК ИНН,
	|	ДанныеСервиса.КПП           КАК КПП
	|ПОМЕСТИТЬ ДанныеСервиса
	|ИЗ
	|	&ДанныеСервиса КАК ДанныеСервиса
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ИНН,
	|	КПП
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ДанныеСервиса.Идентификатор   КАК Идентификатор,
	|	ДанныеСервиса.ИНН             КАК ИНН,
	|	ДанныеСервиса.КПП             КАК КПП,
	|	ТаблицаКОбновлению.Контрагент КАК Контрагент
	|ИЗ
	|	ДанныеСервиса КАК ДанныеСервиса
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ТаблицаКОбновлению КАК ТаблицаКОбновлению
	|		ПО ДанныеСервиса.ИНН = ТаблицаКОбновлению.ИНН
	|			И ДанныеСервиса.КПП = ТаблицаКОбновлению.КПП";
	
	Запрос.УстановитьПараметр("ТаблицаКОбновлению", ТаблицаКОбновлению);
	Запрос.УстановитьПараметр("ДанныеСервиса",      ТаблицаДанныеСервиса);
	
	РезультатЗапроса = Запрос.Выполнить();
	Если РезультатЗапроса.Пустой() Тогда
		Возврат;
	КонецЕсли;
	
	ТаблицаЗаписей = РезультатЗапроса.Выгрузить();
	Отказ       = Ложь;
	ТекстОшибки = "";
	ОбновитьКэшДанныхКонтрагентаБизнесСети(ТаблицаЗаписей, Отказ, ТекстОшибки);
	
	Если Отказ Тогда
		Результат.Успешно     = Ложь;
		Результат.ТекстОшибки = ТекстОшибки;
	Иначе
		Результат.КоличествоЗаписей = ТаблицаЗаписей.Количество();
	КонецЕсли;
	
КонецПроцедуры

Процедура ДобавитьЗаписьТаблицуКонтрагентовПриПолученииИдентификаторов(ТаблицаКОбновлению, Выборка)
	
	НоваяСтрока = ТаблицаКОбновлению.Добавить();
	НоваяСтрока.Контрагент     = Выборка.Контрагент;
	НоваяСтрока.ИНН            = Выборка.ИНН;
	НоваяСтрока.КПП            = Выборка.КПП;
	НоваяСтрока.Идентификатор  = "";
	
КонецПроцедуры

Процедура ДополнитьТаблицуИдентификаторовПоДаннымСервиса(ТаблицаДанныеСервиса, НатуральныеИдентификаторы)
	
	Отказ = Ложь;
	ДанныеСервиса = БизнесСеть.ДанныеОрганизацийПоНатуральнымИдентификаторам(НатуральныеИдентификаторы, Отказ);
	Если Отказ Тогда
		Возврат;
	КонецЕсли;
	
	Для Каждого ЗаписьСервиса Из ДанныеСервиса Цикл
		
		НоваяСтрока = ТаблицаДанныеСервиса.Добавить();
		НоваяСтрока.ИНН           = XMLСтрока(ЗаписьСервиса.ИНН);
		НоваяСтрока.КПП           = XMLСтрока(ЗаписьСервиса.КПП);
		НоваяСтрока.Идентификатор = ЗаписьСервиса.Идентификатор;
		
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#КонецЕсли