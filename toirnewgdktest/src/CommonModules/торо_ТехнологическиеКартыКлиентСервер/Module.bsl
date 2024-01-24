#Область СлужебныйПрограммныйИнтерфейс

#Область ПараллельноеВыполнениеТОТК

Процедура ЗаполнитьИндексКартинкиОперации(СтрокаОперации, ОтображатьПомеченныеНаУдаление = Ложь) Экспорт

	Попытка
		РемонтнаяРабота = СтрокаОперации.РемонтнаяРабота;
	Исключение
		РемонтнаяРабота = СтрокаОперации.Операция;
	КонецПопытки;
	
	Если ОтображатьПомеченныеНаУдаление Тогда
	    ПометкаУдаления = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(РемонтнаяРабота, "ПометкаУдаления");
	КонецЕсли;
	
	Если ОтображатьПомеченныеНаУдаление И ПометкаУдаления Тогда
	    СтрокаОперации.Картинка = ?(ТипЗнч(РемонтнаяРабота) = Тип("СправочникСсылка.торо_ТехнологическиеОперации")
													,?(ЗначениеЗаполнено(СтрокаОперации.ID_ПараллельнойОперации), 8, 3)
													,?(ЗначениеЗаполнено(СтрокаОперации.ID_ПараллельнойОперации), 6, 1));
	Иначе
		СтрокаОперации.Картинка = ?(ТипЗнч(РемонтнаяРабота) = Тип("СправочникСсылка.торо_ТехнологическиеОперации")
													,?(ЗначениеЗаполнено(СтрокаОперации.ID_ПараллельнойОперации), 7, 2)
													,?(ЗначениеЗаполнено(СтрокаОперации.ID_ПараллельнойОперации), 5, 0));
	КонецЕсли;
	
КонецПроцедуры

Процедура ЗаполнитьПараллельноВыполняемыеОперации(СтрокаОперации, ТехКарта) Экспорт

	СтруктураПоиска = Новый Структура("ID_ПараллельнойОперации", СтрокаОперации.ID_ПараллельнойОперации);
	НайденныеСтроки = ТехКарта.СписокОпераций.НайтиСтроки(СтруктураПоиска);
	
	Если НайденныеСтроки.Количество() > 0 Тогда
		ПараллельноВыполняемые = Новый Массив();
		
		Для каждого НайденнаяСтрока Из НайденныеСтроки Цикл
			Если НайденнаяСтрока.ID = СтрокаОперации.ID Или Не ЗначениеЗаполнено(НайденнаяСтрока.ID_ПараллельнойОперации) Тогда
			    Продолжить;
			КонецЕсли;
			
			ПараллельноВыполняемые.Добавить(НайденнаяСтрока.НомерСтроки);
		КонецЦикла;
		
		СтрокаОперации.ПараллельноВыполняемые = СтрСоединить(ПараллельноВыполняемые, ",");
	Иначе
		СтрокаОперации.ПараллельноВыполняемые = "";
	КонецЕсли;
	
КонецПроцедуры

Процедура ЗаполнитьРеквизитыПараллельностиТО(СтрокаОперации, ТехКарта, ОтображатьПомеченныеНаУдаление = Ложь) Экспорт

	ЗаполнитьИндексКартинкиОперации(СтрокаОперации, ОтображатьПомеченныеНаУдаление);
	ЗаполнитьПараллельноВыполняемыеОперации(СтрокаОперации, ТехКарта);
	
КонецПроцедуры

Процедура ЗаполнитьКолонкиПараллельногоВыполненияТО(ДеревоРемонтныхРабот, РемонтныеРаботы, ID_ПараллельнойОперации, СоответствиеРезультата = Неопределено) Экспорт
	
	ТекущиеДанные = ДеревоРемонтныхРабот.ТекущиеДанные;
	
	СтрокаРодитель = ТекущиеДанные.ПолучитьРодителя();
	КоллекцияДляОбхода = СтрокаРодитель.ПолучитьЭлементы();
	
	Для каждого СтрокаДерева Из КоллекцияДляОбхода Цикл
		Если Не ЗначениеЗаполнено(СоответствиеРезультата) = Неопределено И ЗначениеЗаполнено(СоответствиеРезультата[СтрокаДерева.ID]) Тогда
			СтрокаДерева.ID_ПараллельнойОперации = СоответствиеРезультата[СтрокаДерева.ID].ID_ПараллельнойОперации;
		ИначеЕсли Не ЗначениеЗаполнено(СоответствиеРезультата) = Неопределено Тогда
			СтрокаДерева.ID_ПараллельнойОперации = "";
		КонецЕсли;
		
		торо_ТехнологическиеКартыКлиентСервер.ЗаполнитьИндексКартинкиОперации(СтрокаДерева);
		
		// В РМ Диспетчера РР хранятся только в дереве РР, поэтому дополнительно заполнять ТЧ РемонтныеРаботы не нужно
		Если РемонтныеРаботы = Неопределено Тогда
			Продолжить;
		КонецЕсли;
		
		СтруктураПоиска = Новый Структура("ID", СтрокаДерева.ID);
		НайденныеСтроки = РемонтныеРаботы.НайтиСтроки(СтруктураПоиска);
		
		Если Не НайденныеСтроки.Количество() = 0 Тогда
			НайденныеСтроки[0].ID_ПараллельнойОперации = СтрокаДерева.ID_ПараллельнойОперации;
		КонецЕсли;
	КонецЦикла;
	
	// Если в дереве была выбрана не параллельная РР, то после выбора параллельных для нее нужно заполнить
	// реквизит формы ID_ПараллельнойОперации, чтобы условное офомрелние сработало сразу
	Если Не ТекущиеДанные = Неопределено Тогда
		ID_ПараллельнойОперации = ТекущиеДанные.ID_ПараллельнойОперации;
	Иначе
		ID_ПараллельнойОперации = "";
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#КонецОбласти