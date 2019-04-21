using System.Linq;

namespace ViewComPorts
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [System.STAThread]
        static void Main()
        {
            var comp = new NotifyIcon();
            System.Windows.Forms.Application.Run();
        }
    }

}
