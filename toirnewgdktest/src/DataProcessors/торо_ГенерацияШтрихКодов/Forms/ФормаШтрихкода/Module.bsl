
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если ЗначениеЗаполнено(Параметры.ШтрихКод) Тогда 
		ШтрихКод = Параметры.ШтрихКод;
		ТипШтрихкода = Параметры.ТипШтрихКода;
		Если НЕ ЗначениеЗаполнено(ТипШтрихкода) Тогда
			ОбщегоНазначения.СообщитьПользователю(НСтр("ru = 'Тип текущего штрихкода не определен.'"));
		КонецЕсли;
		ОтрисоватьШтрихКод();	
	КонецЕсли; 
	
	Если Параметры.РежимПросмотра Тогда 
		Элементы.Сохранить.Видимость = Ложь;
		Элементы.ФормаСгенерировать.Видимость = Ложь;
		Элементы.ФормаЗакрыть.Видимость = Истина;
		Элементы.ФормаЗакрыть.КнопкаПоУмолчанию = Истина;
		Элементы.ТипШтрихкода.ТолькоПросмотр = Истина;
	КонецЕсли;
	
	Максимумы = Параметры.Максимумы; 
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы 

&НаКлиенте
Процедура ТипШтрихкодаПриИзменении(Элемент)
	ШтрихКод = "";
	СгенерироватьШтрихкодНаСервере();
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура Сгенерировать(Команда)
	Если Не ЗначениеЗаполнено(ТипШтрихкода) Тогда
		ОбщегоНазначенияКлиент.СообщитьПользователю(НСтр("ru = 'Не заполнен тип штрихкода.'"));
	Иначе
		СгенерироватьШтрихкодНаСервере();
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура Сохранить(Команда)
	Если ШтрихКод = "" Тогда 
		Закрыть();
	Иначе
		Закрыть(Новый Структура("ШтрихКод, ТипШтрихКода", ШтрихКод, ТипШтрихкода));
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ЗакрытьФорму(Команда)
	Закрыть();
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура СгенерироватьШтрихкодНаСервере()
	
	Если ЗначениеЗаполнено(Максимумы) Тогда 
		ЛокальныйМаксимум = Макс(Максимумы[торо_ОбщегоНазначенияКлиентСервер.ТипШтрихКодаСтрокой(ТипШтрихкода)], ШтрихКод);
	Иначе
		ЛокальныйМаксимум = ШтрихКод;
	КонецЕсли;
	
	ШтрихКод = Обработки.торо_ГенерацияШтрихКодов.СгенерироватьШтрихкод(ТипШтрихкода, ЛокальныйМаксимум);
	ОтрисоватьШтрихКод();
		
КонецПроцедуры

&НаСервере
Процедура ОтрисоватьШтрихКод()

	ПолеВыводаШтрихКода.Очистить();
	ОбластьВывода = Неопределено;
	Обработки.торо_ГенерацияШтрихКодов.ОтрисоватьШтрихКод(ШтрихКод, ТипШтрихкода, ОбластьВывода);
	ПолеВыводаШтрихКода.Вывести(ОбластьВывода);

КонецПроцедуры

#КонецОбласти
