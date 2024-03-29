#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, СтандартнаяОбработка)
	СтатусДокумента = Перечисления.торо_СтатусыДокументаИнвентаризации.Подготовлен;	
	торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
	Склад = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнойСклад", Истина);					
КонецПроцедуры

Процедура ПриКопировании(ОбъектКопирования)
	
	Дата 		  = НачалоДня(ОбщегоНазначения.ТекущаяДатаПользователя());
	Ответственный = Пользователи.ТекущийПользователь();
	
КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	МассивНепроверяемыхРеквизитов = Новый Массив;
	
	ТипСклада = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Склад, "ТипСклада");
	
	МассивНепроверяемыхРеквизитов.Добавить("Товары.КоличествоУчет");
	МассивНепроверяемыхРеквизитов.Добавить("Товары.Количество");
	
	СтруктураПоиска = Новый Структура("Номенклатура, Характеристика, Серия");

	Для каждого Строка Из Товары Цикл
		
		Префикс = "Товары[" + Формат(Строка.НомерСтроки - 1, "ЧН=0; ЧГ=") + "].";
		ИмяСписка = НСтр("ru = 'Товары'");
		
		Если Строка.Количество = 0 И Строка.КоличествоУчет = 0 Тогда
			
			ИмяПоля = НСтр("ru = 'Количество'");
			Поле = Префикс + "Количество";
			
			ТекстСообщения = ОбщегоНазначенияКлиентСервер.ТекстОшибкиЗаполнения("Колонка", "Заполнение", 
				ИмяПоля, Строка.НомерСтроки, ИмяСписка, ТекстСообщения);
			ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстСообщения, ЭтотОбъект, Поле, "Объект", Отказ);
			
			ИмяПоля = НСтр("ru = 'Количество учетное'");
			Поле = Префикс + "КоличествоУчет";
			
			ТекстСообщения = ОбщегоНазначенияКлиентСервер.ТекстОшибкиЗаполнения("Колонка", "Заполнение", 
				ИмяПоля, Строка.НомерСтроки, ИмяСписка, ТекстСообщения);
			ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстСообщения, ЭтотОбъект, Поле, "Объект", Отказ);
			
		КонецЕсли;
		
		ЗаполнитьЗначенияСвойств(СтруктураПоиска, Строка);
		СтрокиДубли = Товары.НайтиСтроки(СтруктураПоиска);
		Если СтрокиДубли.Количество() > 1 Тогда 
			Отказ = Истина;
			ОбщегоНазначения.СообщитьПользователю(СтрШаблон(НСтр("ru = 'Строка № %1 имеет дубли'"), Строка.НомерСтроки));
		КонецЕсли;
		
	КонецЦикла;
	
	ФОИспользоватьХарактеристикиНоменклатуры = ПолучитьФункциональнуюОпцию("торо_ИспользоватьХарактеристикиНоменклатуры");
	МассивНепроверяемыхРеквизитов.Добавить("Товары.Характеристика");
	Если ФОИспользоватьХарактеристикиНоменклатуры Тогда
		НоменклатураСервер.ПроверитьЗаполнениеХарактеристик(ЭтотОбъект, МассивНепроверяемыхРеквизитов, Отказ);
	КонецЕсли;
	
	ФОИспользоватьСерииНоменклатуры = ПолучитьФункциональнуюОпцию("ИспользоватьСерииНоменклатуры");
	МассивНепроверяемыхРеквизитов.Добавить("Товары.Серия");
	Если ФОИспользоватьСерииНоменклатуры = Истина тогда
		НоменклатураСервер.ПроверитьЗаполнениеСерий(ЭтотОбъект, Отказ, МассивНепроверяемыхРеквизитов);
	КонецЕсли;

	ОбщегоНазначения.УдалитьНепроверяемыеРеквизитыИзМассива(ПроверяемыеРеквизиты, МассивНепроверяемыхРеквизитов);
	
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры


#КонецОбласти

#КонецЕсли
