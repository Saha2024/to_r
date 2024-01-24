#Область ОбработчикиСобытийФормы

&НаКлиенте
&После("Сформировать")
Процедура проф_Сформировать()
	//++ Проф-ИТ, #83, Иванова Е.С., 25.09.2023
	проф_ПроверитьВидОтчетаИДобавитьКоманду();
	//-- Проф-ИТ, #83, Иванова Е.С., 25.09.2023
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

//++ Проф-ИТ, #83, Иванова Е.С., 25.09.2023

&НаКлиенте
Процедура проф_ПечатьГодовойГрафикППРиТО(Команда)
	
	ПараметрыОткрытияФормы = Новый Структура;
	ПараметрыОткрытияФормы.Вставить("ИмяПечатнойФормы", "ИД_ГодовойПланГрафикРемонта");
	ПараметрыОткрытияФормы.Вставить("ПериодОтчета", ПериодОтчета()); 
	
	ОткрытьФорму("Отчет.торо_ПлановыеВнеплановыеРемонты.Форма.проф_ФормаОтборовДляПФ",
		ПараметрыОткрытияФормы, ЭтаФорма, , , , , РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
КонецПроцедуры

&НаКлиенте
Процедура проф_ПечатьМесячныйГрафикППРиТО(Команда)
	
	ПараметрыОткрытияФормы = Новый Структура;
	ПараметрыОткрытияФормы.Вставить("ИмяПечатнойФормы", "ИД_МесячныйПланГрафикРемонта");
	ПараметрыОткрытияФормы.Вставить("ПериодОтчета", ПериодОтчета()); 
	
	ОткрытьФорму("Отчет.торо_ПлановыеВнеплановыеРемонты.Форма.проф_ФормаОтборовДляПФ",
		ПараметрыОткрытияФормы, ЭтаФорма, , , , , РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
КонецПроцедуры

//-- Проф-ИТ, #83, Иванова Е.С., 25.09.2023

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #83, Иванова Е.С., 25.09.2023

// Проверяет Имя, КлючВарианта отчета и его период
// Добавляет команду на форму для необходимого отчета
&НаСервере
Процедура проф_ПроверитьВидОтчетаИДобавитьКоманду()
	
	Если ТипЗнч(НастройкиОтчета.ОтчетСсылка) = Тип("СправочникСсылка.ИдентификаторыОбъектовМетаданных") 
		И НастройкиОтчета.ОтчетСсылка.Имя = "торо_ПлановыеВнеплановыеРемонты" 
		И НастройкиОтчета.ВариантСсылка.КлючВарианта = "ПлановыеРемонты" Тогда
		
		ПериодОтчета = ПериодОтчета();
		
		УдалитьКомандыИЭлементы();
				
		Если ПериодОтчета <> Неопределено
			И ВыбранПериодГод(ПериодОтчета) Тогда
			
			НоваяКоманда = Команды.Добавить("ПечатьГодовойГрафикППРиТО");

			НоваяКоманда.Действие = "проф_ПечатьГодовойГрафикППРиТО";
			НоваяКоманда.Заголовок = "Годовой график ППР и ТО";
			
			НоваяКнопка = Элементы.Вставить("ПечатьГодовойГрафикППРиТО", Тип("КнопкаФормы"), Элементы.ГруппаПечать);
			НоваяКнопка.ИмяКоманды = "ПечатьГодовойГрафикППРиТО";
			НоваяКнопка.Вид = ВидКнопкиФормы.КнопкаКоманднойПанели;
			
		КонецЕсли;
		
		Если ПериодОтчета <> Неопределено
			И ВыбранПериодМесяц(ПериодОтчета) Тогда
			
			НоваяКоманда = Команды.Добавить("ПечатьМесячныйГрафикППРиТО");
			НоваяКоманда.Действие = "проф_ПечатьМесячныйГрафикППРиТО";
			НоваяКоманда.Заголовок = "Месячный ППР и ТО";
			
			НоваяКнопка = Элементы.Вставить("ПечатьМесячныйГрафикППРиТО", Тип("КнопкаФормы"), Элементы.ГруппаПечать);
			НоваяКнопка.ИмяКоманды = "ПечатьМесячныйГрафикППРиТО";
			НоваяКнопка.Вид = ВидКнопкиФормы.КнопкаКоманднойПанели;
			
		КонецЕсли;
	КонецЕсли;

КонецПроцедуры

&НаСервере
Процедура УдалитьКомандыИЭлементы()
	
	КомандаФормы = Команды.Найти("ПечатьГодовойГрафикППРиТО");
	ЭлементФормы = Элементы.Найти("ПечатьГодовойГрафикППРиТО");
	
	Если Не КомандаФормы = Неопределено Тогда
		Команды.Удалить(КомандаФормы);
	КонецЕсли;
	Если Не ЭлементФормы = Неопределено Тогда
		Элементы.Удалить(ЭлементФормы);
	КонецЕсли;
	
	КомандаФормы = Команды.Найти("ПечатьМесячныйГрафикППРиТО");
	ЭлементФормы = Элементы.Найти("ПечатьМесячныйГрафикППРиТО");
	Если Не КомандаФормы = Неопределено Тогда
		Команды.Удалить(КомандаФормы);
	КонецЕсли;
	Если Не ЭлементФормы = Неопределено Тогда
		Элементы.Удалить(ЭлементФормы);
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Функция ПериодОтчета()
	
	Параметр = Отчет.КомпоновщикНастроек.Настройки.ПараметрыДанных.Элементы.Найти("ПериодВыполнения");
	Если Параметр <> Неопределено Тогда
		Элемент = Отчет.КомпоновщикНастроек.ПользовательскиеНастройки.Элементы.Найти(
									  Параметр.ИдентификаторПользовательскойНастройки);
		Если Элемент <> Неопределено Тогда
			Возврат Элемент.Значение;
		КонецЕсли;	
	КонецЕсли;
	
	Возврат Неопределено;
	
КонецФункции

&НаСервереБезКонтекста
Функция ВыбранПериодМесяц(ПериодОтчета)

	НачалоПериодаОтчета = ПериодОтчета.ДатаНачала;
	КонецПериодаОтчета	= ПериодОтчета.ДатаОкончания;
	
	Если НачалоПериодаОтчета = НачалоМесяца(НачалоПериодаОтчета)
		И КонецПериодаОтчета = КонецМесяца(НачалоПериодаОтчета) Тогда
		
		Возврат Истина;
		
	КонецЕсли;
	
	Возврат Ложь;

КонецФункции

&НаСервереБезКонтекста
Функция ВыбранПериодГод(ПериодОтчета)

	НачалоПериодаОтчета = ПериодОтчета.ДатаНачала;
	КонецПериодаОтчета	= ПериодОтчета.ДатаОкончания;
	
	Если НачалоПериодаОтчета = НачалоГода(НачалоПериодаОтчета)
		И КонецПериодаОтчета = КонецГода(КонецПериодаОтчета) Тогда
	
		Возврат Истина;
		
	КонецЕсли;
	
	Возврат Ложь;

КонецФункции 

//-- Проф-ИТ, #83, Иванова Е.С., 25.09.2023

#КонецОбласти
