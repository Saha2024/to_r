
#Область ОбработчикиСобытийФормы

&НаСервере
&После("ПриСозданииНаСервере")
Процедура проф_ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #291, Горетовская М.С., 13.10.2023
	// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"
	ДобавитьЭлементыФормы();
	//-- Проф-ИТ, #291, Горетовская М.С., 13.10.2023
	// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"
	
	//++ Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	проф_ОбщегоНазначенияВызовСервера.УстановитьПараметрВыбораПодразделения(Элементы);
	//-- Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	
КонецПроцедуры

#КонецОбласти

//++ Проф-ИТ, #291, Горетовская М.С., 13.10.2023
// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"

#Область ОбработчикиСобытийЭлементовТаблицыФормыМатериальныеЗатраты
&НаКлиенте
Процедура проф_МатериальныеЗатратыЕдиницаИзмеренияПриИзменении(Элемент)
	
	ТекДанные = Элементы.МатериальныеЗатраты.ТекущиеДанные;
	Если НЕ ТекДанные = Неопределено Тогда
		проф_ПриСменеЗначенийКоличестваИлиУпаковки(ТекДанные);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура проф_МатериальныеЗатратыЕдиницаИзмеренияНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	ТекущиеДанные = Элементы.МатериальныеЗатраты.ТекущиеДанные;
	
	проф_ЗаполнитьДанныеВыбораУпаковки(ТекущиеДанные.Номенклатура, ДанныеВыбора, СтандартнаяОбработка);
	
КонецПроцедуры

#КонецОбласти //ОбработчикиСобытийЭлементовТаблицыФормыМатериальныеЗатраты

#Область ОбработчикиСобытийЭлементовТаблицыФормыСерийныеЗапчасти

&НаКлиенте
Процедура проф_СерийныеЗапчастиЕдиницаИзмеренияНоваяНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	ТекущиеДанные = Элементы.СерийныеЗапчасти.ТекущиеДанные;
	
	проф_ЗаполнитьДанныеВыбораУпаковки(ТекущиеДанные.НоменклатураНовая, ДанныеВыбора, СтандартнаяОбработка);
	
КонецПроцедуры

&НаКлиенте
Процедура проф_СерийныеЗапчастиЕдиницаИзмеренияНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	ТекущиеДанные = Элементы.СерийныеЗапчасти.ТекущиеДанные;
	
	проф_ЗаполнитьДанныеВыбораУпаковки(ТекущиеДанные.Номенклатура, ДанныеВыбора, СтандартнаяОбработка);
	
КонецПроцедуры

#КонецОбласти //ОбработчикиСобытийЭлементовТаблицыФормыСерийныеЗапчасти

#Область ОбработчикиСобытийЭлементовТаблицыФормыДеревоРемонтныхРабот

&НаКлиенте
&ИзменениеИКонтроль("ДеревоРемонтныхРаботПриАктивизацииСтроки")
Процедура проф_ДеревоРемонтныхРаботПриАктивизацииСтроки(Элемент)

	ТекущиеДанные = Элементы.ДеревоРемонтныхРабот.ТекущиеДанные;
	Если Не ТекущиеДанные = Неопределено Тогда
		ID_ПараллельнойОперации = ТекущиеДанные.ID_ПараллельнойОперации;
		#Вставка
		//++ Проф-ИТ, #291, Горетовская М.С., 13.10.2023
		// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"
		проф_УстановитьОтборСтрокПоЗатратам();
		//-- Проф-ИТ, #291, Горетовская М.С., 13.10.2023
		// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"
		#КонецВставки
	Иначе
		ID_ПараллельнойОперации = "";
	КонецЕсли;

КонецПроцедуры

#КонецОбласти //ОбработчикиСобытийЭлементовТаблицыФормыДеревоРемонтныхРабот

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ДобавитьЭлементыФормы()
	
	СтрГруппаФормы = "ГруппаФормы";
	СтрПолеФормы = "ПолеФормы";
	СтрХарактеристика = "Характеристика";
	СтрНачалоВыбора = "НачалоВыбора";
	
	ГруппаРемонтныхРабот = Элементы.Добавить("ГруппаРемонтныхРабот",
											Тип(СтрГруппаФормы),
											Элементы.ГруппаОбследованноеОборудование);
	ГруппаРемонтныхРабот.Вид = ВидГруппыФормы.ОбычнаяГруппа;
	ГруппаРемонтныхРабот.Отображение = ОтображениеОбычнойГруппы.Нет;
	ГруппаРемонтныхРабот.ОтображатьЗаголовок = Ложь; 
	ГруппаРемонтныхРабот.Группировка = ГруппировкаПодчиненныхЭлементовФормы.ГоризонтальнаяВсегда;
	
	Элементы.Переместить(Элементы.ДеревоРемонтныхРабот, ГруппаРемонтныхРабот);
	
	ПанельЗатрат = Элементы.Добавить("ПанельЗатрат", Тип(СтрГруппаФормы), Элементы.ГруппаРемонтныхРабот);
	ПанельЗатрат.Вид = ВидГруппыФормы.Страницы;	
	
	// "Материальные затраты"
	СтраницаМатериальныеЗатраты = Элементы.Добавить("СтраницаМатериальныеЗатраты",
													Тип(СтрГруппаФормы),
													Элементы.ПанельЗатрат);
	СтраницаМатериальныеЗатраты.Вид = ВидГруппыФормы.Страница;
	СтраницаМатериальныеЗатраты.Заголовок = "Материальные затраты";
	СтраницаМатериальныеЗатраты.ПутьКДаннымЗаголовка = "Объект.проф_МатериальныеЗатраты.КоличествоСтрок";
	
	ТаблицаМатериальныеЗатраты = Элементы.Добавить("МатериальныеЗатраты",
													Тип("ТаблицаФормы"),
													Элементы.СтраницаМатериальныеЗатраты);
	ТаблицаМатериальныеЗатраты.ПутьКДанным = "Объект.проф_МатериальныеЗатраты";
	
	ЭлементТаблицы = Элементы.Добавить("МатериальныеЗатратыНоменклатура", Тип(СтрПолеФормы), ТаблицаМатериальныеЗатраты);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_МатериальныеЗатраты.Номенклатура";
	
	ЭлементТаблицы = Элементы.Добавить("МатериальныеЗатратыХарактеристикаНоменклатуры",
										Тип(СтрПолеФормы),
										ТаблицаМатериальныеЗатраты);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_МатериальныеЗатраты.ХарактеристикаНоменклатуры";
	ЭлементТаблицы.Заголовок = СтрХарактеристика;
	
	ЭлементТаблицы = Элементы.Добавить("МатериальныеЗатратыЕдиницаИзмерения",
										Тип(СтрПолеФормы),
										ТаблицаМатериальныеЗатраты);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_МатериальныеЗатраты.ЕдиницаИзмерения";
	ЭлементТаблицы.Заголовок = "Упаковка";
	ЭлементТаблицы.Видимость = Ложь;
	ЭлементТаблицы.УстановитьДействие("ПриИзменении", "проф_МатериальныеЗатратыЕдиницаИзмеренияПриИзменении");
	ЭлементТаблицы.УстановитьДействие(СтрНачалоВыбора, "проф_МатериальныеЗатратыЕдиницаИзмеренияНачалоВыбора");
	
	НоваяСвязь = Новый СвязьПараметраВыбора("Отбор.Номенклатура",
		"Элементы.МатериальныеЗатраты.ТекущиеДанные.Номенклатура");
	
	НовыйМассив = Новый Массив();
	НовыйМассив.Добавить(НоваяСвязь);
	
	НовыеСвязи = Новый ФиксированныйМассив(НовыйМассив);
	ЭлементТаблицы.СвязиПараметровВыбора = НовыеСвязи;
	
	ЭлементТаблицы = Элементы.Добавить("МатериальныеЗатратыНоменклатураЕдиницаИзмерения",
										Тип(СтрПолеФормы),
										ТаблицаМатериальныеЗатраты);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_МатериальныеЗатраты.Номенклатура.ЕдиницаИзмерения";
	ЭлементТаблицы.Заголовок = "Ед. изм.";
	ЭлементТаблицы.УстановитьДействие("ПриИзменении", "проф_МатериальныеЗатратыЕдиницаИзмеренияПриИзменении");
	ЭлементТаблицы.УстановитьДействие(СтрНачалоВыбора, "проф_МатериальныеЗатратыЕдиницаИзмеренияНачалоВыбора");
	
	ЭлементТаблицы = Элементы.Добавить("МатериальныеЗатратыКоличество", Тип(СтрПолеФормы), ТаблицаМатериальныеЗатраты);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_МатериальныеЗатраты.Количество";

	// "Запчасти"
	СтраницаЗапчасти = Элементы.Добавить("СтраницаЗапчасти", Тип(СтрГруппаФормы), Элементы.ПанельЗатрат);
	СтраницаЗапчасти.Вид = ВидГруппыФормы.Страница;
	СтраницаЗапчасти.Заголовок = "Запчасти";
	СтраницаЗапчасти.ПутьКДаннымЗаголовка = "Объект.проф_СерийныеЗапчасти.КоличествоСтрок";
	
	ТаблицаЗапчасти = Элементы.Добавить("СерийныеЗапчасти", Тип("ТаблицаФормы"), Элементы.СтраницаЗапчасти);
	ТаблицаЗапчасти.ПутьКДанным = "Объект.проф_СерийныеЗапчасти";
	
	// Запчасть старая
	ГруппаКолонок = Элементы.Добавить("СерийныеЗапчастиГруппаСтараяЗапчасть", Тип(СтрГруппаФормы), ТаблицаЗапчасти);
	ГруппаКолонок.Вид = ВидГруппыФормы.ГруппаКолонок;
	ГруппаКолонок.Заголовок = "Старая запчасть";
	ГруппаКолонок.Группировка = ГруппировкаКолонок.Горизонтальная;
	ГруппаКолонок.ОтображатьВШапке = Истина;
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиНоменклатура", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.Номенклатура";
	ЭлементТаблицы.Заголовок = "Наименование";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиХарактеристикаНоменклатуры", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.ХарактеристикаНоменклатуры";
	ЭлементТаблицы.Заголовок = СтрХарактеристика;
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиСерияНоменклатуры", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.СерияНоменклатуры";
	ЭлементТаблицы.Заголовок = "Серия";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиСтатусДвиженияСерийныхЗЧ", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.СтатусДвиженияСерийныхЗЧ";
	ЭлементТаблицы.Заголовок = "Статус движения запчасти";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиЕдиницаИзмерения", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.ЕдиницаИзмерения";
	ЭлементТаблицы.Заголовок = "Упаковка, Ед. изм.";
	ЭлементТаблицы.Видимость = Ложь;
	ЭлементТаблицы.УстановитьДействие(СтрНачалоВыбора, "проф_СерийныеЗапчастиЕдиницаИзмеренияНачалоВыбора");
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиКоличество", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.Количество";
	ЭлементТаблицы.Заголовок = "Кол-во";
	
	// Запчасть новая
	ГруппаКолонок = Элементы.Добавить("СерийныеЗапчастиГруппаНоваяЗапчасть", Тип(СтрГруппаФормы), ТаблицаЗапчасти);
	ГруппаКолонок.Вид = ВидГруппыФормы.ГруппаКолонок;
	ГруппаКолонок.Заголовок = "Новая запчасть";
	ГруппаКолонок.Группировка = ГруппировкаКолонок.Горизонтальная;
	ГруппаКолонок.ОтображатьВШапке = Истина;
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиНоменклатураНовая", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.НоменклатураНовая";
	ЭлементТаблицы.Заголовок = "Наименование";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиХарактеристикаНоменклатурыНовая",
										Тип(СтрПолеФормы),
										ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.ХарактеристикаНоменклатурыНовая";
	ЭлементТаблицы.Заголовок = СтрХарактеристика;
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиСерияНоменклатурыНовая", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.СерияНоменклатурыНовая";
	ЭлементТаблицы.Заголовок = "Серия";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиСтатусДвиженияНовойЗЧ", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.СтатусДвиженияНовойЗЧ";
	ЭлементТаблицы.Заголовок = "Статус движения запчасти";
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиЕдиницаИзмеренияНовая", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.ЕдиницаИзмеренияНовая";
	ЭлементТаблицы.Заголовок = "Упаковка, Ед. изм.";
	ЭлементТаблицы.Видимость = Ложь;
	ЭлементТаблицы.УстановитьДействие(СтрНачалоВыбора, "проф_СерийныеЗапчастиЕдиницаИзмеренияНоваяНачалоВыбора");
	
	ЭлементТаблицы = Элементы.Добавить("СерийныеЗапчастиКоличествоНовое", Тип(СтрПолеФормы), ГруппаКолонок);
	ЭлементТаблицы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицы.ПутьКДанным = "Объект.проф_СерийныеЗапчасти.КоличествоНовое";
	ЭлементТаблицы.Заголовок = "Кол-во";
	
	//++ Проф-ИТ, #326, Соловьев А.А., 31.10.2023
	НовыйЭлемент = Элементы.Добавить("проф_ЭксплуатацияВозможна", Тип("ПолеФормы"), Элементы.ШапкаЛево);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеПереключателя;
	НовыйЭлемент.ВидПереключателя = ВидПереключателя.Тумблер;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ЭксплуатацияВозможна";
	НовыйЭлемент.СписокВыбора.Добавить("Да", "Да");
	НовыйЭлемент.СписокВыбора.Добавить("Нет", "Нет");
	//-- Проф-ИТ, #326, Соловьев А.А., 31.10.2023
	
КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Процедура проф_ПриСменеЗначенийКоличестваИлиУпаковки(ТекущиеДанные, ИзменениеКоличестваЕдиниц = Ложь)
	
	КоэффициентУпаковки = проф_ПолучитьКоэффициентУпаковкиНаСервере(ТекущиеДанные.ЕдиницаИзмерения);	
	
	Если КоэффициентУпаковки = 0 Тогда
		КоэффициентУпаковки = 1;
	КонецЕсли; 
	
	Если ИзменениеКоличестваЕдиниц Тогда
		ТекущиеДанные.Количество = ТекущиеДанные.КоличествоЕдиниц / КоэффициентУпаковки;
	Иначе	
		ТекущиеДанные.КоличествоЕдиниц = ТекущиеДанные.Количество * КоэффициентУпаковки;
	КонецЕсли;
	
КонецПроцедуры

&НаСервереБезКонтекста
Функция проф_ПолучитьКоэффициентУпаковкиНаСервере(Упаковка)
	
	Возврат Упаковка.Коэффициент;
	
КонецФункции

&НаКлиенте
Процедура проф_ЗаполнитьДанныеВыбораУпаковки(Номенклатура, ДанныеВыбора, СтандартнаяОбработка)
	
	ПараметрыВыбораУпаковки = Новый Структура("Номенклатура", Номенклатура);
	СтандартнаяОбработка = Ложь;
	проф_ЗаполнитьДанныеВыбораУпаковкиСервер(ДанныеВыбора, ПараметрыВыбораУпаковки);
	
КонецПроцедуры

&НаСервереБезКонтекста
Процедура проф_ЗаполнитьДанныеВыбораУпаковкиСервер(ДанныеВыбора, ПараметрыВыбора)
	
	ДанныеВыбора = Справочники.УпаковкиНоменклатуры.ПолучитьДанныеВыбора(ПараметрыВыбора);
	
КонецПроцедуры

&НаКлиенте
// Процедура устанавливает отбор строк табличных полей затрат по строке  дерева ремонтных работ.
//
Процедура проф_УстановитьОтборСтрокПоЗатратам()
	
	ТекСтрокаДереваРемонтов = Элементы.ДеревоРемонтныхРабот.ТекущиеДанные;
	ТекСтрокаОбследованноеОборудование = Элементы.ОбследованноеОборудование.ТекущиеДанные;
	
	ID_Отбора = ТекСтрокаДереваРемонтов.ID;
	
	Если ТекСтрокаОбследованноеОборудование <> Неопределено Тогда
		ID_ОтбораРемонтов = ТекСтрокаОбследованноеОборудование.ID;
		Элементы.МатериальныеЗатраты.ОтборСтрок = Новый ФиксированнаяСтруктура("ID, РемонтыОборудования_ID",
																				ID_Отбора, ID_ОтбораРемонтов);
	Иначе
		Элементы.МатериальныеЗатраты.ОтборСтрок = Новый ФиксированнаяСтруктура("ID", ID_Отбора);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти //СлужебныеПроцедурыИФункции

//-- Проф-ИТ, #291, Горетовская М.С., 13.10.2023
// "Невыполненные ремонтные работы" при формировании документа "Смета (заявка на ремонт)"
