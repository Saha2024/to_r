#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
Перем БезусловнаяЗапись Экспорт; // Отключает проверки при записи документа
перем СтруктураДанных Экспорт;  // Структура, хранящая данные для работы с уведомлениями.
	
#Область ОбработчикиСобытий

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли; 
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, Режим)
	
	Движения.торо_ЦелевыеЗначенияKPI.Записывать = Истина;
	Движения.торо_ЦелевыеЗначенияKPI.Очистить();
	
	Для Каждого Строка Из КлючевыеПоказатели Цикл
		
		Движение = Движения.торо_ЦелевыеЗначенияKPI.Добавить();
		
		ЗаполнитьЗначенияСвойств(Движение, Строка);

		Движение.Период = Дата;
		Движение.ДатаПо = ?(Строка.ЗначениеОграниченоПоВремени, Строка.ДатаПО, '39991231235959');
	КонецЦикла;

КонецПроцедуры

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, СтандартнаяОбработка)
	
	Если ТипЗнч(ДанныеЗаполнения) = Тип("Структура") Тогда
		ЗаполнитьЗначенияСвойств(ЭтотОбъект, ДанныеЗаполнения);
	КонецЕсли;
	
	Если НЕ ЗначениеЗаполнено(Организация) Тогда
		Организация = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнаяОрганизация");
	КонецЕсли;
	
	Если НЕ ЗначениеЗаполнено(Ответственный) Тогда
		Ответственный = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнойОтветственный");
		Если НЕ ЗначениеЗаполнено(Ответственный) Тогда
			Ответственный = Пользователи.ТекущийПользователь();
		КонецЕсли;
	КонецЕсли;

КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	СообщенияОбОшибках = Новый Соответствие; // Ключ - Текст сообщения; Значение - путь к данным.
	
	Показатели = ОбщегоНазначения.ВыгрузитьКолонку(КлючевыеПоказатели, "Показатель", Истина);
	ЦелевыеТрендыПоказателей = ОбщегоНазначения.ЗначениеРеквизитаОбъектов(Показатели, "ЦелевойТренд");
	
	Для Каждого Строка ИЗ КлючевыеПоказатели Цикл
		
		ЦелевойТренд = ЦелевыеТрендыПоказателей.Получить(Строка.Показатель);
		
		Если Строка.ЗначениеОграниченоПоВремени И Строка.ДатаПо < Дата Тогда
			
			ТекстСообщения = СтрШаблон(НСтр("ru = 'Дата действия целевых значений показателя в строке %1 не может быть меньше даты документа.'"), Строка.НомерСтроки);
			ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].ДатаПо", Строка.НомерСтроки - 1);
			СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
		КонецЕсли;
		
		Если Не ЗначениеЗаполнено(ЦелевойТренд) Тогда
			
			ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 не заполнен целевой тренд.'"), Строка.НомерСтроки);
			ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].ЦелевойТрендПоказателя", Строка.НомерСтроки - 1);
			СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
			
		ИначеЕсли ЦелевойТренд = Перечисления.торо_ВидыЦелевыхТрендовПоказателейKPI.Минимизация 
			ИЛИ ЦелевойТренд = Перечисления.торо_ВидыЦелевыхТрендовПоказателейKPI.Максимизация Тогда
			
			Если Не ЗначениеЗаполнено(Строка.Значение) Тогда  
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 не заполнено значение.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].Значение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным); 
				
			ИначеЕсли ЦелевойТренд = Перечисления.торо_ВидыЦелевыхТрендовПоказателейKPI.Минимизация
				И ЗначениеЗаполнено(Строка.КритическоеЗначение) И Строка.КритическоеЗначение < Строка.Значение Тогда
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 критическое значение не может быть меньше значения.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].КритическоеЗначение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным); 
				
			ИначеЕсли ЦелевойТренд = Перечисления.торо_ВидыЦелевыхТрендовПоказателейKPI.Максимизация
				И Строка.КритическоеЗначение > Строка.Значение Тогда
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 критическое значение не может быть больше значения.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].КритическоеЗначение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);	
			КонецЕсли;	
			
		ИначеЕсли ЦелевойТренд = Перечисления.торо_ВидыЦелевыхТрендовПоказателейKPI.Диапазон Тогда
			
			Если Не ЗначениеЗаполнено(Строка.МинимальноеЗначение) Тогда 
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 не заполнено минимальное значение.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].МинимальноеЗначение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
				
			ИначеЕсли Строка.КритическийМинимум > Строка.МинимальноеЗначение Тогда
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 критический минимум не может быть больше минимального значения.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].КритическийМинимум", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
			КонецЕсли;
			
			Если Не ЗначениеЗаполнено(Строка.МаксимальноеЗначение) Тогда 
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 не заполнено максимальное значение.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].МаксимальноеЗначение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);  
				
			ИначеЕсли ЗначениеЗаполнено(Строка.КритическийМаксимум) И Строка.КритическийМаксимум < Строка.МаксимальноеЗначение Тогда
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 критический максимум не может быть меньше максимального значения.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].КритическийМаксимум", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
			КонецЕсли;
			
			Если ЗначениеЗаполнено(Строка.МинимальноеЗначение) И ЗначениеЗаполнено(Строка.МаксимальноеЗначение)
				И Строка.МинимальноеЗначение >= Строка.МаксимальноеЗначение Тогда
				
				ТекстСообщения = СтрШаблон(НСтр("ru = 'Для показателя в строке %1 максимум должен быть больше минимума.'"), Строка.НомерСтроки);
				ПутьКДанным = СтрШаблон("Объект.КлючевыеПоказатели[%1].МаксимальноеЗначение", Строка.НомерСтроки - 1);
				СообщенияОбОшибках.Вставить(ТекстСообщения, ПутьКДанным);
			КонецЕсли;
		КонецЕсли;	
	КонецЦикла;	
	
	Если СообщенияОбОшибках.Количество() > 0 Тогда
		
		ТекстСообщения = НСтр("ru = 'При проверке корректности заполнения произошли ошибки. Проведение невозможно!'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);	
		
		Для Каждого КлючИЗначение Из СообщенияОбОшибках Цикл 
			ОбщегоНазначения.СообщитьПользователю(КлючИЗначение.Ключ,, КлючИЗначение.Значение);
		КонецЦикла; 
		
		Отказ = Истина;
	КонецЕсли;	

КонецПроцедуры

#КонецОбласти

#КонецЕсли