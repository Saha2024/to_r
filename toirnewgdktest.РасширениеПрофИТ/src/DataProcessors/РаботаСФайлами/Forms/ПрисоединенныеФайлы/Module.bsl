
&НаКлиенте
&ИзменениеИКонтроль("ДобавитьФайл")
Процедура проф_ДобавитьФайл()

	ТекущиеДанные = Элементы.Список.ТекущиеДанные;
	ГруппаФайлов = Неопределено;
	Если ТекущиеДанные <> Неопределено И ТекущиеДанные.ЭтоГруппа И ТекущиеДанные.ВладелецФайла = Параметры.ВладелецФайла Тогда
		ГруппаФайлов = ТекущиеДанные.Ссылка;
	КонецЕсли;
	Если ЭтоВладелецЭлементовСправочникаФайлы Тогда
		РаботаСФайламиСлужебныйКлиент.ДобавитьФайлИзФайловойСистемы(Параметры.ВладелецФайла, ЭтотОбъект);
	Иначе
		#Вставка
		//++ Проф-ИТ, #157, Иванова Е.С., 22.09.2023
		проф_ПроверитьВозможностьДобавленияФайла(Параметры.ВладелецФайла);
		//-- Проф-ИТ, #157, Иванова Е.С., 22.09.2023
        #КонецВставки
		РаботаСФайламиКлиент.ДобавитьФайлы(Параметры.ВладелецФайла, УникальныйИдентификатор, , ГруппаФайлов);
	КонецЕсли;

КонецПроцедуры

//++ Проф-ИТ, #157, Иванова Е.С., 22.09.2023
&НаСервереБезКонтекста
Процедура проф_ПроверитьВозможностьДобавленияФайла(Знач ВладелецФайла)
	
	ТипЗначения = ТипЗнч(ВладелецФайла);
	ОбъектМД 	=  Метаданные.НайтиПоТипу(ТипЗначения);
	Если Метаданные.Справочники.Содержит(ОбъектМД) Тогда
		Если (ЗначениеЗаполнено(ВладелецФайла) И ВладелецФайла.ЭтоГруппа)
			Или (Не ЗначениеЗаполнено(ВладелецФайла) 
			И ТипЗнч(ВладелецФайла) = Тип("СправочникСсылка.торо_ТехКарты")) Тогда
				Шаблон = НСтр("ru = 'К группе справочников нельзя присоединять файлы'");
				ВызватьИсключение Шаблон;
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры 
//-- Проф-ИТ, #157, Иванова Е.С., 22.09.2023

&НаКлиенте
Процедура проф_ДобавитьФайлПоШаблонуВместо(Команда)
	
	ПараметрыДобавления = Новый Структура;
	ПараметрыДобавления.Вставить("ОбработчикРезультата",          Неопределено);
	ПараметрыДобавления.Вставить("ВладелецФайла",                 ВладелецФайла);
	ПараметрыДобавления.Вставить("ФормаВладелец",                 ЭтотОбъект);
	ПараметрыДобавления.Вставить("ИмяСправочникаХранилищаФайлов", ИмяСправочникаХранилищаФайлов);
	
	//++ Проф-ИТ, #157, Иванова Е.С., 22.09.2023
	проф_ПроверитьВозможностьДобавленияФайла(ВладелецФайла);
	//-- Проф-ИТ, #157, Иванова Е.С., 22.09.2023
	
	РаботаСФайламиСлужебныйКлиент.ДобавитьНаОсновеШаблона(ПараметрыДобавления);
	

КонецПроцедуры
