#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #157, Башинская А.Ю., 24.09.2023    
	Если Элементы.Найти("Форматоро_Перейти") <> Неопределено 
			И Элементы.Форматоро_Перейти.ПодчиненныеЭлементы.Найти("ФормаОбработкаторо_РаботаСОбщимиПрисоединеннымиФайламиОбщиеПрисоединенныеФайлы") <> Неопределено Тогда     
		Элементы.Форматоро_Перейти.ПодчиненныеЭлементы.ФормаОбработкаторо_РаботаСОбщимиПрисоединеннымиФайламиОбщиеПрисоединенныеФайлы.Видимость = Ложь;
	КонецЕсли;
	//-- Проф-ИТ, #157, Башинская А.Ю., 24.09.2023
	
КонецПроцедуры

#КонецОбласти