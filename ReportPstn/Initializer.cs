using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportPstn
{
    class Initializer
    {
        internal static int[] GetMvzGts()
        {
            int[] mvzGts =
            {
                0,         530100000, 530200000, 530202000, 530204020,
                530300000, 530300010, 530300040, 530600000, 530600010,
                530600020, 530600030, 530600040, 530601000, 530700000,
                530700010, 530700020, 530700030, 530700040, 530700070,
                530701000, 530701010, 530702000, 530800000, 530800010,
                530800020, 530800030, 530800050, 530800060, 530800070,
                530800080, 530900000, 530900010, 530900020, 530900030,
                530901000, 531100010, 531100020, 531100030, 531100050,
                531100060, 531101020, 531100070, 531200000, 531200010,
                531200020, 531200040, 531200050, 531300010, 531300020,
                531300100, 531700000, 531701000, 531702000, 531703000,
                531800040, 531900000, 531910000, 531910020, 530700090,
            };

            return mvzGts;
        }

        internal static int[] GetMvzRtk()
        {
            int[] mvzRtk =
            {
                0,         530100000, 530200000, 530202000, 530300000,
                530300010, 530300040, 530600000, 530600010, 530600020,
                530600030, 530600040, 530601000, 530700000, 530700010,
                530700020, 530700030, 530700040, 530700070, 530701000,
                530701010, 530702000, 530800000, 530800010, 530800020,
                530800030, 530800050, 530800060, 530800070, 530800080,
                530900000, 530900010, 530900020, 530900030, 530901000,
                530902000, 531100000, 531100010, 531100020, 531100030,
                531100060, 531101020, 531100070, 531200000, 531200010,
                531200020, 531200040, 531200050, 531300010, 531300020,
                531700000, 531701000, 531702000, 531703000, 531800040,
                531900000, 531910000, 531910020, 530700090,
            };

            return mvzRtk;
        }

        internal static string[] GetMvzCaption()
        {
            string[] caption = { "МВЗ", "Заказ", "Сумма по МВЗ" };
            return caption;
        }

        internal static string[] GetStatisticCaption()
        {
            string[] caption = { "Дата", "Время", "Город", "Абонент", "Мин", "Сумма", "Код", "Выз_телефон", "МВЗ", "Заказ" };
            return caption;
        }

        internal static string[] GetMvzPhoneCaption()
        {
            string[] caption = { "№ телефона", "Наименование заказа" };
            return caption;
        }

        internal static Dictionary<string, string> GetMvzPhoneDictionary()
        {
            Dictionary<string, string> mvzPhone = new Dictionary<string, string>();
            mvzPhone.Add("20266", "531200000"); mvzPhone.Add("20356", "531200020");
            mvzPhone.Add("20600", "531200000"); mvzPhone.Add("20673", "0");
            mvzPhone.Add("21101", "531200000"); mvzPhone.Add("21529", "531200000");
            mvzPhone.Add("21956", "531300050"); mvzPhone.Add("22054", "530600000");
            mvzPhone.Add("22800", "531300040"); mvzPhone.Add("23122", "530702000");
            mvzPhone.Add("23133", "531100010"); mvzPhone.Add("23144", "531100010");
            mvzPhone.Add("23510", "531300050"); mvzPhone.Add("24377", "531300100");
            mvzPhone.Add("26092", "530800000"); mvzPhone.Add("26138", "530702000");
            mvzPhone.Add("26488", "531200000"); mvzPhone.Add("62267", "531200040");
            mvzPhone.Add("63445", "531100050"); mvzPhone.Add("63644", "531200000");
            mvzPhone.Add("73004", "530600030"); mvzPhone.Add("73006", "530600010");
            mvzPhone.Add("73027", "530600010"); mvzPhone.Add("73038", "530800030");
            mvzPhone.Add("73040", "530900020"); mvzPhone.Add("73050", "531300020");
            mvzPhone.Add("73107", "531100010"); mvzPhone.Add("73167", "530600030");
            mvzPhone.Add("73180", "531100000"); mvzPhone.Add("73237", "530600020");
            mvzPhone.Add("73238", "530700070"); mvzPhone.Add("73263", "530800030");
            mvzPhone.Add("73267", "530600010"); mvzPhone.Add("73287", "530600000");
            mvzPhone.Add("73317", "530600010"); mvzPhone.Add("73318", "530600040");
            mvzPhone.Add("73320", "530800030"); mvzPhone.Add("73321", "530600010");
            mvzPhone.Add("73347", "531200020"); mvzPhone.Add("73368", "531200010");
            mvzPhone.Add("73372", "530600000"); mvzPhone.Add("73380", "531200050");
            mvzPhone.Add("73441", "530700000"); mvzPhone.Add("73456", "530100000");
            mvzPhone.Add("73512", "530600010"); mvzPhone.Add("73531", "531100020");
            mvzPhone.Add("73570", "531910000"); mvzPhone.Add("73571", "531910000");
            mvzPhone.Add("73582", "531910000"); mvzPhone.Add("73602", "530800030");
            mvzPhone.Add("73611", "531200010"); mvzPhone.Add("73666", "530200000");
            mvzPhone.Add("73680", "531200000"); mvzPhone.Add("73684", "531100010");
            mvzPhone.Add("73704", "530600010"); mvzPhone.Add("73790", "531200040");
            mvzPhone.Add("73801", "530800000"); mvzPhone.Add("73810", "530800000");
            mvzPhone.Add("73820", "");          mvzPhone.Add("73822", "530600010");
            mvzPhone.Add("73828", "530600010"); mvzPhone.Add("73831", "530600010");
            mvzPhone.Add("73860", "531100010"); mvzPhone.Add("73872", "531200000");
            mvzPhone.Add("73881", "");          mvzPhone.Add("73901", "531100030");
            mvzPhone.Add("73903", "530700090"); mvzPhone.Add("73905", "531200000");
            mvzPhone.Add("73923", "530600000"); mvzPhone.Add("73950", "531100010");
            mvzPhone.Add("73961", "530601000"); mvzPhone.Add("73966", "530600040");
            mvzPhone.Add("77090", "530300000"); mvzPhone.Add("77117", "530900030");
            mvzPhone.Add("77121", "531100010"); mvzPhone.Add("77154", "530600040");
            mvzPhone.Add("77168", "530800010"); mvzPhone.Add("77197", "530800030");
            mvzPhone.Add("77269", "530800050"); mvzPhone.Add("77296", "530601000");
            mvzPhone.Add("77331", "530800030"); mvzPhone.Add("77372", "530300000");
            mvzPhone.Add("77401", "531200000"); mvzPhone.Add("77405", "531200040");
            mvzPhone.Add("77423", "531200010"); mvzPhone.Add("77431", "531200000");
            mvzPhone.Add("77465", "530202000"); mvzPhone.Add("77517", "530600010");
            mvzPhone.Add("77531", "531100010"); mvzPhone.Add("77552", "530700000");
            mvzPhone.Add("77555", "530600010"); mvzPhone.Add("77710", "531200000");
            mvzPhone.Add("77765", "530700030"); mvzPhone.Add("77830", "530300000");
            mvzPhone.Add("77843", "531200000"); mvzPhone.Add("77856", "530900000");
            mvzPhone.Add("77989", "531200000"); mvzPhone.Add("77988", "530800000");
            mvzPhone.Add("77991", "530700000"); mvzPhone.Add("77996", "531200010");
            mvzPhone.Add("79925", "530200000"); mvzPhone.Add("79932", "531200000");
            mvzPhone.Add("79933", "531200000"); mvzPhone.Add("79945", "530701000");

            return mvzPhone;
        }

        internal static string[] GetMvzOrderCaption()
        {
            string[] caption = { "Наименование заказа", "№ заказа" };
            return caption;
        }

        internal static Dictionary<string, string> GetMvzOrderDictionary()
        {
            Dictionary<string, string> mvzOrder = new Dictionary<string, string>();
            mvzOrder.Add("0", ""); mvzOrder.Add("530100000", "538090000013");
            mvzOrder.Add("530100010", "538090000015"); mvzOrder.Add("530100020", "538090000014");
            mvzOrder.Add("530200000", "538090000016"); mvzOrder.Add("530202000", "538090000017");
            mvzOrder.Add("530202010", "538090000018"); mvzOrder.Add("530203010", "538090000019");
            mvzOrder.Add("530300000", "538090000020"); mvzOrder.Add("530300010", "538090000021");
            mvzOrder.Add("530300020", "538090000022"); mvzOrder.Add("530300040", "538090000023");
            mvzOrder.Add("530301010", "538090000024"); mvzOrder.Add("530600000", "538090000025");
            mvzOrder.Add("530600010", "538090000030"); mvzOrder.Add("530600030", "538090000027");
            mvzOrder.Add("530600040", "538090000029"); mvzOrder.Add("530600050", "538090000026");
            mvzOrder.Add("530600060", "538090000028"); mvzOrder.Add("530601000", "");
            mvzOrder.Add("530700000", "538090000031"); mvzOrder.Add("530700020", "538090000032");
            mvzOrder.Add("530700030", "538090000033"); mvzOrder.Add("530700040", "538090000034");
            mvzOrder.Add("530700060", "538090000035"); mvzOrder.Add("530700090", "");
            mvzOrder.Add("530701000", "538090000036"); mvzOrder.Add("530701020", "538090000037");
            mvzOrder.Add("530800000", "538090000038"); mvzOrder.Add("530800010", "538090000042");
            mvzOrder.Add("530800020", "538090000041"); mvzOrder.Add("530800030", "538090000054");
            mvzOrder.Add("530800050", "");             mvzOrder.Add("530800070", "538090000039");
            mvzOrder.Add("530800090", "538090000040"); mvzOrder.Add("530900020", "538090000043");
            mvzOrder.Add("530900030", "538090000043"); mvzOrder.Add("531100000", "538090000044");
            mvzOrder.Add("531100010", "538090000045"); mvzOrder.Add("531100020", "538090000046");
            mvzOrder.Add("531100030", "");             mvzOrder.Add("531100060", "");
            mvzOrder.Add("531100070", "");             mvzOrder.Add("531101020", "");
            mvzOrder.Add("531200000", "538090000047"); mvzOrder.Add("531200010", "538090000048");
            mvzOrder.Add("531200020", "");             mvzOrder.Add("531200040", "538090000049");
            mvzOrder.Add("531200050", "");             mvzOrder.Add("531300020", "");
            mvzOrder.Add("531800000", "538090000050"); mvzOrder.Add("531800020", "538090000051");
            mvzOrder.Add("531800030", "538090000052"); mvzOrder.Add("531910000", "");
            mvzOrder.Add("531910010", "538090000053"); mvzOrder.Add("999999999", "");

            return mvzOrder;
        }

        internal static string[] GetCorporateCaption()
        {
            string[] caption = { "№ телефона", "Категория" };
            return caption;
        }

        internal static Dictionary<string, string> GetCorporateDictionary()
        {
            Dictionary<string, string> corporate = new Dictionary<string, string>();
            corporate.Add("73004", "4"); corporate.Add("73006", "4");
            corporate.Add("73027", "2"); corporate.Add("73040", "2");
            corporate.Add("73050", "2"); corporate.Add("73107", "2");
            corporate.Add("73167", "2"); corporate.Add("73180", "9");
            corporate.Add("73237", "2"); corporate.Add("73238", "2");
            corporate.Add("73263", "9"); corporate.Add("73267", "4");
            corporate.Add("73287", "4"); corporate.Add("73317", "4");
            corporate.Add("73318", "4"); corporate.Add("73320", "2");
            corporate.Add("73321", "4"); corporate.Add("73347", "16");
            corporate.Add("73368", "2"); corporate.Add("73372", "4");
            corporate.Add("73380", "6"); corporate.Add("73441", "2");
            corporate.Add("73456", "4"); corporate.Add("73512", "4");
            corporate.Add("73531", "9"); corporate.Add("73570", "2");
            corporate.Add("73571", "2"); corporate.Add("73582", "6");
            corporate.Add("73602", "2"); corporate.Add("73611", "2");
            corporate.Add("73666", "4"); corporate.Add("73680", "4");
            corporate.Add("73684", "4"); corporate.Add("73704", "4");
            corporate.Add("73790", "2"); corporate.Add("73801", "2");
            corporate.Add("73810", "2"); corporate.Add("73820", "4");
            corporate.Add("73822", "4"); corporate.Add("73828", "4");
            corporate.Add("73831", "4"); corporate.Add("73860", "4");
            corporate.Add("73872", "2"); corporate.Add("73881", "2");
            corporate.Add("73901", "2"); corporate.Add("73903", "2");
            corporate.Add("73905", "4"); corporate.Add("73923", "4");
            corporate.Add("73950", "4"); corporate.Add("73961", "2");
            corporate.Add("73966", "4"); corporate.Add("77090", "2");
            corporate.Add("77117", "2"); corporate.Add("77121", "6");
            corporate.Add("77154", "2"); corporate.Add("77168", "2");
            corporate.Add("77197", "2"); corporate.Add("77269", "2");
            corporate.Add("77296", "2"); corporate.Add("77331", "9");
            corporate.Add("77372", "4"); corporate.Add("77401", "4");
            corporate.Add("77405", "2"); corporate.Add("77423", "2");
            corporate.Add("77431", "4"); corporate.Add("77465", "4");
            corporate.Add("77517", "4"); corporate.Add("77531", "2");
            corporate.Add("77552", "4"); corporate.Add("77555", "4");
            corporate.Add("77710", "4"); corporate.Add("77765", "2");
            corporate.Add("77830", "2"); corporate.Add("77843", "4");
            corporate.Add("77856", "4"); corporate.Add("77988", "2");
            corporate.Add("77991", "4"); corporate.Add("77996", "2");
            corporate.Add("79925", "2"); corporate.Add("79932", "4");
            corporate.Add("79933", "4"); corporate.Add("79945", "2");

            return corporate;
        }
    }
}
