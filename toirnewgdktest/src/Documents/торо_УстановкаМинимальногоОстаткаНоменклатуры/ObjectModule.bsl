#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
	
////////////////////////////////////////////////////////////////////////////////
// СЛУЖЕБНЫЕ ПЕРЕМЕННЫЕ

перем СтруктураДанных Экспорт;  // Структура, хранящая данные для работы с уведомлениями.
Перем БезусловнаяЗапись Экспорт; // Отключает проверки при записи документа

#Область ОбработчикиСобытий

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, ТекстЗаполнения, СтандартнаяОбработка)
	
	Если ТипЗнч(ДанныеЗаполнения) = Тип("Структура") Тогда
		ЗаполнитьЗначенияСвойств(ЭтотОбъект, ДанныеЗаполнения);
	КонецЕсли;

	Если НЕ ЗначениеЗаполнено(Организация) Тогда
		Организация = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнаяОрганизация");
	КонецЕсли;
	
	Если НЕ ЗначениеЗаполнено(Склад) Тогда
		Склад = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнойСклад");
	КонецЕсли;
	
КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	СтруктураПоиска = Новый Структура("Номенклатура, ХарактеристикаНоменклатуры");
	Для Каждого Строка Из СписокНоменклатуры Цикл 
		ЗаполнитьЗначенияСвойств(СтруктураПоиска, Строка);
		СтрокиДубли = СписокНоменклатуры.НайтиСтроки(СтруктураПоиска);
		Если СтрокиДубли.Количество() > 1 Тогда 
			Отказ = Истина;
			ОбщегоНазначения.СообщитьПользователю(СтрШаблон(НСтр("ru = 'Строка № %1 имеет дубли'"), Строка.НомерСтроки));
		КонецЕсли;
	КонецЦикла;

	МассивНепроверяемыхРеквизитов = Новый Массив;
	МассивВсехРеквизитов = Новый Массив;
	МассивРеквизитовОперации = Новый Массив;
	
	ОбщегоНазначенияУТКлиентСервер.ЗаполнитьМассивНепроверяемыхРеквизитов(
		МассивВсехРеквизитов,
		МассивРеквизитовОперации,
		МассивНепроверяемыхРеквизитов);
		
	ФОИспользоватьХарактеристикиНоменклатуры = ПолучитьФункциональнуюОпцию("торо_ИспользоватьХарактеристикиНоменклатуры");
	МассивНепроверяемыхРеквизитов.Добавить("Номенклатура.ХарактеристикаНоменклатуры");
	Если ФОИспользоватьХарактеристикиНоменклатуры = Истина Тогда
		ПараметрыПроверки = Новый Структура("СуффиксХарактеристики, ИмяТЧ", "Номенклатуры", "СписокНоменклатуры");
		НоменклатураСервер.ПроверитьЗаполнениеХарактеристик(ЭтотОбъект,МассивНепроверяемыхРеквизитов, Отказ, ПараметрыПроверки);
	КонецЕсли;
	
	Если ОбщегоНазначенияУТ.ПроверитьЗаполнениеРеквизитовОбъекта(ЭтотОбъект, ПроверяемыеРеквизиты) Тогда
		Отказ = Истина;
	КонецЕсли;
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, РежимПроведения)
	
	Движения.торо_МинимальныеОстаткиНоменклатуры.Записывать = Истина;
	
	Для каждого СтрокаНоменклатуры Из СписокНоменклатуры Цикл
	
		Движение = Движения.торо_МинимальныеОстаткиНоменклатуры.Добавить();
		
		Движение.Период = ДатаУстановки;
		Движение.Организация = Организация;
		Движение.Склад = Склад;
		Движение.Номенклатура = СтрокаНоменклатуры.Номенклатура;
		Движение.ХарактеристикаНоменклатуры = СтрокаНоменклатуры.ХарактеристикаНоменклатуры;
		Если ТипЗнч(СтрокаНоменклатуры.ЕдиницаИзмерения) = Тип("СправочникСсылка.УпаковкиНоменклатуры") Тогда
			Движение.МинимальныйОстаток = СтрокаНоменклатуры.МинимальныйОстаток * ?(СтрокаНоменклатуры.ЕдиницаИзмерения.Коэффициент = 0, 1, СтрокаНоменклатуры.ЕдиницаИзмерения.Коэффициент);
		Иначе
			Движение.МинимальныйОстаток = СтрокаНоменклатуры.МинимальныйОстаток;
		КонецЕсли; 	
	КонецЦикла; 
	
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры

#КонецОбласти

#КонецЕсли