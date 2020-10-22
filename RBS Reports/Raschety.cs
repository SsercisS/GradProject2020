using System;


namespace RBS_Reports
{
    class Raschety
    {
        #region Этапы

        
        /// <summary>
        /// Функция рассчета показателя для 00. Не рассмотренно
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string NeRassmotreno(string data)
        {
            string date = data;
            date = date.Remove(10);
            
            DateTime dateAfter = Convert.ToDateTime(date);
            dateAfter= dateAfter.AddDays(1);

            date = dateAfter.ToShortDateString();
            return date;
        }

        /// <summary>
        /// Функция расчета показателя для 02. Подготовка ИС/ИДОЗ/ЗапросТКП
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="Ozydanye"></param>
        /// <param name="Nerassmotrena"></param>
        /// <returns></returns>
        public static string Podgotovka(string TypeOfTender, string Ozydanye, string Nerassmotrena)
        {
            string date = Nerassmotrena;
           

            DateTime dateAfter = Convert.ToDateTime(date);

            switch (TypeOfTender)
            {
                case "Тендер":

                    if (Ozydanye==null | Ozydanye == "")
                    {
                        dateAfter = dateAfter.AddDays(5);
                        date = dateAfter.ToShortDateString();
                    }
                    else
                    {
                        dateAfter = dateAfter.AddDays(6);
                        date = dateAfter.ToShortDateString();
                    }

                    break;
                case "АР":

                    if (Ozydanye == null | Ozydanye == "")
                    {
                        date = dateAfter.ToShortDateString();
                    }
                    else
                    {
                        dateAfter = dateAfter.AddDays(1);
                        date = dateAfter.ToShortDateString();
                    }

                    break;
                default:
                    if (Ozydanye == null | Ozydanye == "")
                    {
                        dateAfter = dateAfter.AddDays(1);
                        date = dateAfter.ToShortDateString();
                    }
                    else
                    {
                        dateAfter = dateAfter.AddDays(2);
                        date = dateAfter.ToShortDateString();
                    }
                    break;
            }
            
            return date;
        }
        /// <summary>
        /// Функция расчета показателя для 03. Согласование ИДОЗ/ЗапросТКП
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="podgotovka"></param>
        /// <returns></returns>
        public static string Soglasovanie(string TypeOfTender, string podgotovka)
        {
            string date = podgotovka;

            DateTime dateAfter = Convert.ToDateTime(date);

            switch (TypeOfTender)
            {
                case "АР":
                    date = null;
                    break;
                case "МЗ":
                    dateAfter = dateAfter.AddDays(1);
                    date = dateAfter.ToShortDateString();
                    break;
                default:
                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;
            }

                    return date;
        }
        /// <summary>
        /// Функция расчета показателя для 04. Сбор ТКП
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="podgotovka"></param>
        /// <param name="otrabotka"></param>
        /// <param name="soglasovanie"></param>
        /// <returns></returns>
        public static string Sbor(string TypeOfTender, string podgotovka, string otrabotka, string soglasovanie)
        {
            string date = null;

            switch (TypeOfTender)
            {
                case "АР":

                    date = podgotovka;
                    
                    DateTime dateAfter = Convert.ToDateTime(date);

                    dateAfter = dateAfter.AddDays(5);
                    date = dateAfter.ToShortDateString();

                    break;
                case "МЗ":
                    
                    if (otrabotka== null | otrabotka== " ")
                    {
                        date = soglasovanie;

                        dateAfter = Convert.ToDateTime(date);

                        dateAfter = dateAfter.AddDays(5);
                        date = dateAfter.ToShortDateString();
                    }
                    else
                    {
                        date = otrabotka;
                        
                        dateAfter = Convert.ToDateTime(date);
                        dateAfter = dateAfter.AddDays(5);
                        date = dateAfter.ToShortDateString();

                    }

                    break;
                default:

                    date = soglasovanie;

                    dateAfter = Convert.ToDateTime(date);

                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;
            }

            return date;
        }

        /// <summary>
        /// Функция расчета показателя для 05. Формирование АС/оценка тендера
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="sbor"></param>
        /// <returns></returns>
        public static string Formirovanye(string TypeOfTender, string sbor)
        {
            string date = sbor;
            
            DateTime dateAfter = Convert.ToDateTime(date);

            switch (TypeOfTender)
            {
                case "АР":
                    dateAfter = dateAfter.AddDays(1);
                    date = dateAfter.ToShortDateString();
                    break;

                case "МЗ":
                    dateAfter = dateAfter.AddDays(2);
                    date = dateAfter.ToShortDateString();
                    break;

                default:
                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;
            }

            return date;
        }
        /// <summary>
        /// Функция расчета показателя для 06. Согласование итогов (ИС/АС)
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="formirovanie"></param>
        /// <returns></returns>
        public static string SoglasovanieItogy(string TypeOfTender, string formirovanie)
        {
            string date = formirovanie;

            DateTime dateAfter = Convert.ToDateTime(date);

            switch (TypeOfTender)
            {
                case "АР":
                    
                    break;

                case "МЗ":
                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;

                default:
                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;
            }

            return date;
        }
        /// <summary>
        /// Функция расчета показателя для 07. Передано в ЗК/ЦЗК/СЗ
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="itogy"></param>
        /// <param name="formirovanie"></param>
        /// <param name="otrabotka"></param>
        /// <returns></returns>
        public static string Peredano(string TypeOfTender, string itogy, string formirovanie, string otrabotka)
        {
            string date = " ";

            switch (TypeOfTender)
            {
                default:
                    date = formirovanie;

                    DateTime dateAfter = Convert.ToDateTime(date);

                    dateAfter = dateAfter.AddDays(3);
                    date = dateAfter.ToShortDateString();
                    break;

                case "АР":

                    break;

                case "МЗ":
                    if (otrabotka== null | otrabotka== " ")
                    {
                        date = itogy;

                        dateAfter = Convert.ToDateTime(date);

                        dateAfter = dateAfter.AddDays(1);
                        date = dateAfter.ToShortDateString();

                    }
                    else
                    {
                        date = otrabotka;

                        dateAfter = Convert.ToDateTime(date);

                        dateAfter = dateAfter.AddDays(1);
                        date = dateAfter.ToShortDateString();
                    }
                    break;                
            }

            return date;
        }

        #endregion

        #region Реестр заявок
        /// <summary>
        /// Функция расчета показателя для Окончание срока план
        /// </summary>
        /// <param name="TypeOfTender"></param>
        /// <param name="raspred"></param>
        /// <returns></returns>
        public static string OkonchanyePlan(string TypeOfTender, string raspred)
        {
            string date = raspred;
            date = date.Remove(10);
            DateTime dateAfter;

            switch (TypeOfTender)
            {
                case "МЗ":
                    dateAfter = Convert.ToDateTime(date);
                    dateAfter = dateAfter.AddDays(10);
                    date = dateAfter.ToShortDateString();
                    break;
                case "АР":
                    dateAfter = Convert.ToDateTime(date);
                    dateAfter = dateAfter.AddDays(5);
                    date = dateAfter.ToShortDateString();
                    break;
                case "ЕП":
                    dateAfter = Convert.ToDateTime(date);
                    dateAfter = dateAfter.AddDays(16);
                    date = dateAfter.ToShortDateString();
                    break;
                default:
                    dateAfter = Convert.ToDateTime(date);
                    dateAfter = dateAfter.AddDays(31);
                    date = dateAfter.ToShortDateString();
                    break;
            }

            return date;
        }
        /// <summary>
        /// Функция расчета показателя для Фактическое окончание
        /// </summary>
        /// <param name="okonchPlan"></param>
        /// <param name="dneyProdlenia"></param>
        /// <returns></returns>
        public static string OkonchanyeFakt(string okonchPlan, string dneyProdlenia)
        {
            string date = okonchPlan;

            DateTime dateAfter;
            dateAfter = Convert.ToDateTime(date);
            dateAfter = dateAfter.AddDays(Convert.ToInt32(dneyProdlenia));
            date = dateAfter.ToShortDateString();
            return date;
        }
        /// <summary>
        /// Функция расчета показателя для Просроченные заявки
        /// </summary>
        /// <param name="datePeredachi"></param>
        /// <param name="okonchFakt"></param>
        /// <returns></returns>
        public static string Prosrochka(string datePeredachi, string okonchFakt)
        {
            string date = okonchFakt;
            DateTime dateAfter;
            DateTime dateDef;
            
            switch (datePeredachi)
            {
                case "":
                    dateAfter = Convert.ToDateTime(date);
                    if (DateTime.Now.CompareTo(dateAfter)>0)
                    {
                        date = "Просрочено";
                    }
                    else
                    {
                        date = "-";
                    }

                    break;
                default:
                    dateDef = Convert.ToDateTime(datePeredachi);
                    dateAfter = Convert.ToDateTime(okonchFakt);

                    if (dateDef.CompareTo(dateAfter) > 0)
                    {
                        date = "Просрочено";
                    }
                    else
                    {
                        date = "-";
                    }
                    break;
            }

            return date;
        }
        /// <summary>
        /// Функция расчета показателя для Дней просрочки
        /// </summary>
        /// <param name="datePeredachi"></param>
        /// <param name="okonchFakt"></param>
        /// <returns></returns>
        public static int DneyProsrochky(string datePeredachi, string okonchFakt)
        {
            string date = okonchFakt;
            DateTime dateAfter;
            DateTime dateDef;

            int days=0;

            switch (datePeredachi)
            {
                case "":
                    dateAfter = Convert.ToDateTime(date);
                    

                    if (DateTime.Now.CompareTo(dateAfter) > 0)
                    {
                        if (datePeredachi == "")
                        {
                            days = DateTime.Now.CompareTo(dateAfter);
                        }
                        else
                        {
                            dateDef = Convert.ToDateTime(datePeredachi);
                            days = dateDef.CompareTo(dateAfter);
                        }
                    }
                        break;
                default:
                    dateAfter = Convert.ToDateTime(date);
                    dateDef = Convert.ToDateTime(datePeredachi);
                    if (dateDef.CompareTo(dateAfter) > 0)
                    {
                        if (datePeredachi == "")
                        {
                            days = DateTime.Now.CompareTo(dateAfter);
                        }
                        else
                        {
                            days = dateDef.CompareTo(dateAfter);
                        }
                    }
                        break;
            }
            return days;
        }
        /// <summary>
        /// Функция расчета показателя для Категория штрафов
        /// </summary>
        /// <param name="dneyProsrochki"></param>
        /// <returns></returns>
        public static string Shtrafy(string dneyProsrochki)
        {
            string date = dneyProsrochki;

            if (Convert.ToInt32(dneyProsrochki) == 0)
            {
                date = "0";
            }
            else if (Convert.ToInt32(dneyProsrochki) < 6)
            {
                date = "Категория 1";
            }
            else if (Convert.ToInt32(dneyProsrochki) < 11)
            {
                date = "Категория 2";
            }
            else if (Convert.ToInt32(dneyProsrochki) > 10)
            {
                date = "Категория 3";
            }

            
            return date;
        }

        #endregion
    }
}
