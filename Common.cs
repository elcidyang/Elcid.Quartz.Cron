using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Elcid.Quartz.Cron
{
    /// <summary>
    /// 公共类库
    /// </summary>
    public class Common
    {
        public static bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }
    }
}
