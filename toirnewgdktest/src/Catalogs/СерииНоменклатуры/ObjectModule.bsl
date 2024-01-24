#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	НепроверямыеРеквизиты = Новый Массив;
	
	Если ВидНоменклатуры.НастройкаИспользованияСерий <> Перечисления.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеруИСрокуГодности
		И ВидНоменклатуры.НастройкаИспользованияСерий <> Перечисления.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоСрокуГодности Тогда
		НепроверямыеРеквизиты.Добавить("ГоденДо");
	КонецЕсли;

	Если ВидНоменклатуры.НастройкаИспользованияСерий <> Перечисления.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеруИСрокуГодности
		И ВидНоменклатуры.НастройкаИспользованияСерий <> Перечисления.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеру
		И ВидНоменклатуры.НастройкаИспользованияСерий <> Перечисления.НастройкиИспользованияСерийНоменклатуры.ЭкземплярТовара
		Тогда
		НепроверямыеРеквизиты.Добавить("Номер");
	КонецЕсли;
	
	ОбщегоНазначения.УдалитьНепроверяемыеРеквизитыИзМассива(ПроверяемыеРеквизиты, НепроверямыеРеквизиты);
	
КонецПроцедуры

#КонецОбласти
#КонецЕсли