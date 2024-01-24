
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #367, Соловьев А.А., 22.11.2023
	Список.ПроизвольныйЗапрос = Истина;
	СтруктураСвойствСписка = ОбщегоНазначения.СтруктураСвойствДинамическогоСписка();
	СтруктураСвойствСписка.ТекстЗапроса = ТекстЗапросаСписок();
	ОбщегоНазначения.УстановитьСвойстваДинамическогоСписка(Элементы.Список, СтруктураСвойствСписка);
	
	НоваяКолонкаТаблицы = Элементы.Добавить("ЕстьФайлы", Тип("ПолеФормы"), Элементы.Список);
	НоваяКолонкаТаблицы.ПутьКДанным = "Список.ЕстьФайлы"; 
	НоваяКолонкаТаблицы.Вид = ВидПоляФормы.ПолеКартинки;
	НоваяКолонкаТаблицы.Видимость = Истина;
	НоваяКолонкаТаблицы.КартинкаЗначений = БиблиотекаКартинок.Скрепка;
	НоваяКолонкаТаблицы.КартинкаШапки = БиблиотекаКартинок.Скрепка;
	НоваяКолонкаТаблицы.ПоложениеЗаголовка = ПоложениеЗаголовкаЭлементаФормы.Нет;
	
	Элементы.Переместить(НоваяКолонкаТаблицы, Элементы.Список, Элементы.Ссылка);
	//-- Проф-ИТ, #367, Соловьев А.А., 22.11.2023
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #367, Соловьев А.А., 22.11.2023

&НаСервере
Функция ТекстЗапросаСписок()
	
	Возврат
	"ВЫБРАТЬ
	|	ВнутреннееПотреблениеТоваров.Ссылка КАК Ссылка,
	|	ВнутреннееПотреблениеТоваров.ВерсияДанных КАК ВерсияДанных,
	|	ВнутреннееПотреблениеТоваров.ПометкаУдаления КАК ПометкаУдаления,
	|	ВнутреннееПотреблениеТоваров.Номер КАК Номер,
	|	ВнутреннееПотреблениеТоваров.Дата КАК Дата,
	|	ВнутреннееПотреблениеТоваров.Проведен КАК Проведен,
	|	ВнутреннееПотреблениеТоваров.Организация КАК Организация,
	|	ВнутреннееПотреблениеТоваров.Склад КАК Склад,
	|	ВнутреннееПотреблениеТоваров.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	ВнутреннееПотреблениеТоваров.Ответственный КАК Ответственный,
	|	ВнутреннееПотреблениеТоваров.Комментарий КАК Комментарий,
	|	ВнутреннееПотреблениеТоваров.ДатаРаспоряжения КАК ДатаРаспоряжения,
	|	ВнутреннееПотреблениеТоваров.ПотреблениеПоЗаказам КАК ПотреблениеПоЗаказам,
	|	ВнутреннееПотреблениеТоваров.ЕстьРасхождения КАК ЕстьРасхождения,
	|	ВнутреннееПотреблениеТоваров.Подразделение КАК Подразделение,
	|	ВнутреннееПотреблениеТоваров.торо_Автор КАК торо_Автор,
	|	ВнутреннееПотреблениеТоваров.проф_НаправлениеДеятельности КАК проф_НаправлениеДеятельности,
	|	ВнутреннееПотреблениеТоваров.проф_ДокументОснование КАК проф_ДокументОснование,
	|	ВнутреннееПотреблениеТоваров.проф_ЗаказДляЗакупки КАК проф_ЗаказДляЗакупки,
	|	ВнутреннееПотреблениеТоваров.проф_ХозяйственнаяОперация КАК проф_ХозяйственнаяОперация,
	|	ВнутреннееПотреблениеТоваров.МоментВремени КАК МоментВремени,
	|	ВЫБОР
	|		КОГДА НаличиеФайлов.ЕстьФайлы ЕСТЬ NULL
	|			ТОГДА 1
	|		КОГДА НаличиеФайлов.ЕстьФайлы
	|			ТОГДА 0
	|		ИНАЧЕ 1
	|	КОНЕЦ КАК ЕстьФайлы
	|ИЗ
	|	Документ.ВнутреннееПотреблениеТоваров КАК ВнутреннееПотреблениеТоваров
	|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.НаличиеФайлов КАК НаличиеФайлов
	|		ПО ВнутреннееПотреблениеТоваров.Ссылка = НаличиеФайлов.ОбъектСФайлами";
	
КонецФункции

//-- Проф-ИТ, #367, Соловьев А.А., 22.11.2023

#КонецОбласти