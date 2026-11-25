using System;

namespace ExcelBan
{
    public static class JalaliDate
    {
        private static readonly int[] GDaysInMonth = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        private static readonly int[] JDaysInMonth = { 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29 };

        public static (int Year, int Month, int Day) FromGregorian(int gy, int gm, int gd)
        {
            int gy2 = gy - 1600;
            int gm2 = gm - 1;
            int gd2 = gd - 1;

            long gDayNo = 365L * gy2 + (gy2 + 3) / 4 - (gy2 + 99) / 100 + (gy2 + 399) / 400;
            for (int i = 0; i < gm2; i++)
                gDayNo += GDaysInMonth[i];
            if (gm2 > 1 && ((gy % 4 == 0 && gy % 100 != 0) || gy % 400 == 0))
                gDayNo += 1;
            gDayNo += gd2;

            long jDayNo = gDayNo - 79;

            long jNp = jDayNo / 12053;
            jDayNo %= 12053;

            long jy = 979 + 33 * jNp + 4 * (jDayNo / 1461);
            jDayNo %= 1461;

            if (jDayNo >= 366)
            {
                jy += (jDayNo - 1) / 365;
                jDayNo = (jDayNo - 1) % 365;
            }

            int jm = 12;
            long jd = jDayNo + 1;
            for (int i = 0; i < 11; i++)
            {
                if (jDayNo < JDaysInMonth[i])
                {
                    jm = i + 1;
                    jd = jDayNo + 1;
                    break;
                }
                jDayNo -= JDaysInMonth[i];
            }

            return ((int)jy, jm, (int)jd);
        }

        public static string TodayYyMmDd()
        {
            var today = DateTime.Now;
            var (jy, jm, jd) = FromGregorian(today.Year, today.Month, today.Day);
            return $"{jy % 100:D2}{jm:D2}{jd:D2}";
        }
    }
}
