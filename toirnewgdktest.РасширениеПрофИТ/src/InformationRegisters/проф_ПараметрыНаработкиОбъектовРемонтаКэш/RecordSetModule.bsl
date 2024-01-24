#Область ОбработчикиСобытий

&После("ПриЗаписи")
Процедура проф_ПриЗаписи(Отказ, Замещение)

	// ++ Проф-ИТ, #150, Соловьев А., 13.10.2023
	Если ЗначениеЗаполнено(ЭтотОбъект.Отбор.ОбъектРемонта.Значение) 
	И ТипЗнч(ЭтотОбъект.Отбор.ОбъектРемонта.Значение) = Тип("СправочникСсылка.торо_ОбъектыРемонта") 
	И Не ЭтотОбъект.ДополнительныеСвойства.Свойство("НеОбновлятьРегистрНаличиеНормативов") Тогда
		торо_РаботаСНормативамиСервер.ОбновитьСвойствоОР_ЕстьПараметрыНаработки(ЭтотОбъект.Отбор.ОбъектРемонта.Значение);
	КонецЕсли;	
	// -- Проф-ИТ, #150, Соловьев А., 13.10.2023
	
КонецПроцедуры

#КонецОбласти
