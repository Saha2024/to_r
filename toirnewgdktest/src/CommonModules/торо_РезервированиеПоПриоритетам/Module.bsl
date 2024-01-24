#Область ПрограммныйИнтерфейс
// Процедура обновляет нормированное значение критичности дефектов.
&НаСервере
Процедура ОбновитьНормированноеЗначениеКритичностиДефектов() Экспорт
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	МАКСИМУМ(торо_КритичностьДефекта.Порядок) КАК Значение
	               |ИЗ
	               |	Справочник.торо_КритичностьДефекта КАК торо_КритичностьДефекта
	               |ГДЕ
	               |	НЕ торо_КритичностьДефекта.ПометкаУдаления
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_КритичностьДефекта.Ссылка КАК Ссылка,
	               |	торо_КритичностьДефекта.Порядок КАК Значение
	               |ИЗ
	               |	Справочник.торо_КритичностьДефекта КАК торо_КритичностьДефекта
	               |ГДЕ
	               |	НЕ торо_КритичностьДефекта.ПометкаУдаления";
	РезЗапроса = Запрос.ВыполнитьПакет();
	Если НЕ РезЗапроса[1].Пустой() Тогда
		ВыборкаМаксимума = РезЗапроса[0].Выбрать();
		ВыборкаМаксимума.Следующий();
		Максимум =  ?(ВыборкаМаксимума.Значение = 0, 1, ВыборкаМаксимума.Значение);

		Выборка = РезЗапроса[1].Выбрать();
		Пока Выборка.Следующий() Цикл
			СпрОбъект = Выборка.Ссылка.ПолучитьОбъект();
			СпрОбъект.НормированноеЗначение = Выборка.Значение / Максимум;
			СпрОбъект.Записать();
		КонецЦикла;
	КонецЕсли;
	
КонецПроцедуры

// Процедура обновляет нормированное значение критичности объекта ремонта.
&НаСервере
Процедура ОбновитьНормированноеЗначениеКритичностиОР() Экспорт
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	МАКСИМУМ(торо_ПриоритетыОбъектовРемонта.Значение) КАК Значение
	               |ИЗ
	               |	Справочник.торо_ПриоритетыОбъектовРемонта КАК торо_ПриоритетыОбъектовРемонта
	               |ГДЕ
	               |	НЕ торо_ПриоритетыОбъектовРемонта.ПометкаУдаления
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_ПриоритетыОбъектовРемонта.Ссылка КАК Ссылка,
	               |	торо_ПриоритетыОбъектовРемонта.Значение КАК Значение
	               |ИЗ
	               |	Справочник.торо_ПриоритетыОбъектовРемонта КАК торо_ПриоритетыОбъектовРемонта
	               |ГДЕ
	               |	НЕ торо_ПриоритетыОбъектовРемонта.ПометкаУдаления";
	РезЗапроса = Запрос.ВыполнитьПакет();
	Если НЕ РезЗапроса[1].Пустой() Тогда
		ВыборкаМаксимума = РезЗапроса[0].Выбрать();
		ВыборкаМаксимума.Следующий();
		Максимум =  ?(ВыборкаМаксимума.Значение = 0, 1, ВыборкаМаксимума.Значение);

		Выборка = РезЗапроса[1].Выбрать();

		Пока Выборка.Следующий() Цикл
			СпрОбъект = Выборка.Ссылка.ПолучитьОбъект();
			СпрОбъект.НормированноеЗначение = Выборка.Значение / Максимум;
			СпрОбъект.Записать();
		КонецЦикла;
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти



