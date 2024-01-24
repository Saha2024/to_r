
#Область ОбработчикиСобытий

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	Если НЕ ИспользоватьПометкуДанных Тогда
		ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти("ТекстЗапросаПометкиОбрабатываемых"));
	КонецЕсли;
	
	Если НЕ ИспользоватьУдалениеДанных Тогда
		ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти("ТекстЗапросаУдаленияОбработанных"));
	КонецЕсли;
	
	Если ИспользоватьПроизвольнуюСтрокуПодключения Тогда
	    ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти("Сервер"));
		ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти("БазаДанных"));
	Иначе
		ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти("ПроизвольнаяСтрокаПодключения"));
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти