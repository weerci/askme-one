using System;
using System.Collections.Generic;
using System.Text;

namespace Askme
{
    /// <summary>
    /// Проверяет наступление условий для начала проведения работ по формированию запроса, получения
    /// информации из базы и отсылки полученных сведений по e-mail/
    /// </summary>
    class CheckSendCondition
    {
        /// <summary>
        /// Проверка наступления условия для начала формирования нового запроса к базе и отправке e-mail
        /// </summary>
        /// <returns>Проверяется время и состояние отправки.</returns>
        public bool CheckCondition()
        {
            return CheckTime() && CheckSended();
        }

        // Private methods
        private bool CheckTime()
        {
            long currTick = DateTime.Now.Ticks;

            return currTick >= Convert.ToDateTime(Common.SysProp.FromTime).Ticks && 
                currTick <= Convert.ToDateTime(Common.SysProp.ToTime).Ticks;
        }
        private bool CheckSended()
        {
            return Common.SysProp.LastSendedDate != DateTime.Now.ToShortDateString();
        }
    }
}
