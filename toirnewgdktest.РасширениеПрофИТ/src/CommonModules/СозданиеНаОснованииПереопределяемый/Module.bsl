
&После("ПриОпределенииОбъектовСКомандамиСозданияНаОсновании")
Процедура проф_ПриОпределенииОбъектовСКомандамиСозданияНаОсновании(Объекты)
	
	// ++ Проф-ИТ, #313, Корнилов М.С., 04.11.2023
	Объекты.Добавить(Метаданные.Документы.ЗаказНаВнутреннееПотребление);
	// -- Проф-ИТ, #313, Корнилов М.С., 04.11.2023
	// ++ Проф-ИТ, #314, Корнилов М.С., 04.11.2023
	Объекты.Добавить(Метаданные.Документы.проф_ЗаказНаПеремещение);	
	// -- Проф-ИТ, #314, Корнилов М.С., 04.11.2023
	
КонецПроцедуры
