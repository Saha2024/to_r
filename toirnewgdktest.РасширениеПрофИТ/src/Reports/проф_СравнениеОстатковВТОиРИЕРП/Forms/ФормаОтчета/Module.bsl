//++ Проф-ИТ, #398, Соловьев А.А., 18.12.2023

#Область ОбработчикиСобытийФормы

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	ПараметрыПоУмолчанию = ПараметрыПоУмолчанию();
	Склады = ПараметрыПоУмолчанию.Склады;
	Период = ПараметрыПоУмолчанию.Период;
	
	Для Каждого Настройка Из Отчет.КомпоновщикНастроек.ПользовательскиеНастройки.Элементы Цикл
		
		Если Строка(Настройка.Параметр) = "Период" Тогда 
			Настройка.Значение = Период;
			Настройка.Использование = Истина;
		КонецЕсли;
		
		Если ЗначениеЗаполнено(Склады) И Строка(Настройка.Параметр) = "Склад" Тогда  
			Настройка.Значение = Склады;
			Настройка.Использование = Истина;
		КонецЕсли;
		
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервереБезКонтекста
Функция ПараметрыПоУмолчанию()
	
	СтурктураВозврата = Новый Структура;
	
	КодУзла = Справочники.проф_НастройкиСистемы.ПолучитьНастройкуСистемы("НастройкиОбменаТОИР", "КодУзлаЕРП");
	
	Узел = ПланыОбмена.ОбменТОИР30ЕРП20.НайтиПоКоду(КодУзла);
	Если ЗначениеЗаполнено(Узел) Тогда 
		
		Склады = Узел.Склады.ВыгрузитьКолонку("Склад");
		СписокСкладов = Новый СписокЗначений;
		Для Каждого Склад Из Склады Цикл
			СписокСкладов.Добавить(Склад);
		КонецЦикла;
		
	Иначе
		СписокСкладов = Неопределено;
	КонецЕсли;
	
	СтурктураВозврата.Вставить("Период", ТекущаяДатаСеанса());
	СтурктураВозврата.Вставить("Склады", СписокСкладов);
	
	Возврат СтурктураВозврата;
	
КонецФункции

#КонецОбласти

//-- Проф-ИТ, #398, Соловьев А.А., 18.12.2023