      
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПеред(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #4, Башинская А.Ю., 20.08.2023
	СписокРазрешений.ТекстЗапроса = проф_ПолучитьТекстЗапроса();
	
	//Добавление колонок списка
	СписокРазрешений.КомпоновщикНастроек.Восстановить(СпособВосстановленияНастроекКомпоновкиДанных.Полное);
	   
	НовЭлемент = Элементы.Добавить("проф_Статус", Тип("ПолеФормы"), Элементы.Список);
	НовЭлемент.Вид	 = ВидПоляФормы.ПолеНадписи;
	НовЭлемент.ПутьКДанным	= "СписокРазрешений.проф_Статус";	
	НовЭлемент.Заголовок	= "Статус";
	//-- Проф-ИТ, #4, Башинская А.Ю., 20.08.2023

КонецПроцедуры

#КонецОбласти      

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #4, Башинская А.Ю., 20.08.2023

&НаСервере
Функция проф_ПолучитьТекстЗапроса()
	
	Возврат  "ВЫБРАТЬ
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Ссылка КАК Ссылка,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ВерсияДанных КАК ВерсияДанных,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ПометкаУдаления КАК ПометкаУдаления,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Номер КАК Номер,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Дата КАК Дата,
	|	ВЫБОР
	|		КОГДА торо_УстановкаАналоговНоменклатурыДляРемонтов.СписокОбъектовРМ = ЗНАЧЕНИЕ(Справочник.торо_СписокОбъектовРегламентногоМероприятия.ПустаяСсылка)
	|			ТОГДА торо_УстановкаАналоговНоменклатурыДляРемонтов.ОбъектРемонта
	|		ИНАЧЕ торо_УстановкаАналоговНоменклатурыДляРемонтов.СписокОбъектовРМ
	|	КОНЕЦ КАК ОбъектРемонта,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Проведен КАК Проведен,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Организация КАК Организация,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ТехКарта КАК ТехКарта,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.НаправлениеОбъектаРемонта КАК НаправлениеОбъектаРемонта,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ВидРемонта КАК ВидРемонта,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ДатаНачалаДействия КАК ДатаНачалаДействия,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.ДатаОкончанияДействия КАК ДатаОкончанияДействия,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.УказаниеПоПрименению КАК УказаниеПоПрименению,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Ответственный КАК Ответственный,
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.Подразделение КАК Подразделение, 
	|	торо_УстановкаАналоговНоменклатурыДляРемонтов.проф_Статус КАК проф_Статус, 
	|	ВЫБОР
	|		КОГДА НЕ ТаблицаМатериалы.Ссылка ЕСТЬ NULL
	|			ТОГДА ЗНАЧЕНИЕ(Перечисление.ИспользованиеНоменклатурыВНСИПроизводства.Материал)
	|		КОГДА НЕ ТаблицаАналоги.Ссылка ЕСТЬ NULL
	|			ТОГДА ЗНАЧЕНИЕ(Перечисление.ИспользованиеНоменклатурыВНСИПроизводства.Аналог)
	|	КОНЕЦ КАК ИспользуетсяКак
	|ИЗ
	|	Документ.торо_УстановкаАналоговНоменклатурыДляРемонтов КАК торо_УстановкаАналоговНоменклатурыДляРемонтов
	|		ЛЕВОЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
	|			торо_УстановкаАналоговНоменклатурыДляРемонтовМатериалы.Ссылка КАК Ссылка
	|		ИЗ
	|			Документ.торо_УстановкаАналоговНоменклатурыДляРемонтов.Материалы КАК торо_УстановкаАналоговНоменклатурыДляРемонтовМатериалы
	|		ГДЕ
	|			торо_УстановкаАналоговНоменклатурыДляРемонтовМатериалы.Номенклатура = &РазрешенияНоменклатура) КАК ТаблицаМатериалы
	|		ПО торо_УстановкаАналоговНоменклатурыДляРемонтов.Ссылка = ТаблицаМатериалы.Ссылка
	|		ЛЕВОЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
	|			торо_УстановкаАналоговНоменклатурыДляРемонтовАналоги.Ссылка КАК Ссылка
	|		ИЗ
	|			Документ.торо_УстановкаАналоговНоменклатурыДляРемонтов.Аналоги КАК торо_УстановкаАналоговНоменклатурыДляРемонтовАналоги
	|		ГДЕ
	|			торо_УстановкаАналоговНоменклатурыДляРемонтовАналоги.Номенклатура = &РазрешенияНоменклатура) КАК ТаблицаАналоги
	|		ПО торо_УстановкаАналоговНоменклатурыДляРемонтов.Ссылка = ТаблицаАналоги.Ссылка
	|ГДЕ
	|	ВЫБОР
	|			КОГДА &ОтборПоНоменклатуре
	|				ТОГДА ВЫБОР
	|						КОГДА НЕ ТаблицаМатериалы.Ссылка ЕСТЬ NULL
	|								ИЛИ НЕ ТаблицаАналоги.Ссылка ЕСТЬ NULL
	|							ТОГДА ИСТИНА
	|						ИНАЧЕ ЛОЖЬ
	|					КОНЕЦ
	|			ИНАЧЕ ИСТИНА
	|		КОНЕЦ
	|	И ВЫБОР
	|			КОГДА НЕ &ЗначениеФО
	|				ТОГДА торо_УстановкаАналоговНоменклатурыДляРемонтов.СписокОбъектовРМ = ЗНАЧЕНИЕ(Справочник.торо_СписокОбъектовРегламентногоМероприятия.ПустаяСсылка)
	|			ИНАЧЕ ИСТИНА
	|		КОНЕЦ";
	
КонецФункции

//-- Проф-ИТ, #4, Башинская А.Ю., 20.08.2023

#КонецОбласти 