#Область ОбработчикиСобытийФормы
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если Не Параметры.Свойство("ДокументППР") Тогда
		ТекстСообщения = НСтр("ru = 'Для открытия формы необходимо указать документ план-график ППР'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,, Отказ);
		Возврат;
	КонецЕсли;
	
	МассивВыбранныхВидовРемонтов = Неопределено;
	Если Параметры.Свойство("МассивВыбранныхВидовРемонтов") Тогда
	    МассивВыбранныхВидовРемонтов = Параметры.МассивВыбранныхВидовРемонтов;
	КонецЕсли;
	
	Если Не Параметры.Свойство("МассивВидовРемонтов") Тогда
	    Возврат;
	КонецЕсли;
	
	Для каждого ВидРемонта Из Параметры.МассивВидовРемонтов Цикл
	    НоваяСтрока = ВидыРемонтов.Добавить();
		НоваяСТрока.ВидРемонта = ВидРемонта;
		
		Если МассивВыбранныхВидовРемонтов = Неопределено Тогда
		    НоваяСтрока.Пометка = Истина;
		ИначеЕсли МассивВыбранныхВидовРемонтов.Найти(ВидРемонта) = Неопределено Тогда
			НоваяСтрока.Пометка = Ложь;
		Иначе
			НоваяСтрока.Пометка = Истина;
		КонецЕсли;
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы
&НаКлиенте
Процедура УстановитьФлажки(Команда)
	
	Для каждого ВидРемонта Из ВидыРемонтов Цикл
	    ВидРемонта.Пометка = Истина;
	КонецЦикла; 
	
КонецПроцедуры

&НаКлиенте
Процедура СнятьФлажки(Команда)
	
	Для каждого ВидРемонта Из ВидыРемонтов Цикл
	    ВидРемонта.Пометка = Ложь;
	КонецЦикла; 
	
КонецПроцедуры

&НаКлиенте
Процедура ИнвертироватьФлажки(Команда)
	
	Для каждого ВидРемонта Из ВидыРемонтов Цикл
	    ВидРемонта.Пометка = Не ВидРемонта.Пометка;
	КонецЦикла; 
	
КонецПроцедуры

&НаКлиенте
Процедура Выбрать(Команда)
	
	МассивВыбранныхВидовРемонтов = Новый Массив;
	Для каждого ВидРемонта Из ВидыРемонтов Цикл
	    Если ВидРемонта.Пометка Тогда
			МассивВыбранныхВидовРемонтов.Добавить(ВидРемонта.ВидРемонта);
		КонецЕсли; 
	КонецЦикла;
	
	Закрыть(МассивВыбранныхВидовРемонтов);
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции


#КонецОбласти