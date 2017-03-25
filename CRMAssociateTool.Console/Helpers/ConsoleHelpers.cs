using System;
using System.Threading;
using System.Threading.Tasks;

namespace CrmDeployTool.Console.Helpers
{
    public static class ConsoleHelpers
    {
        private static int _counter;
        private static bool _run;

        public static void AnimateOnce()
        {
            _counter++;
            switch (_counter % 4)
            {
                case 0: System.Console.Write("/"); _counter = 0; break;
                case 1: System.Console.Write("-"); break;
                case 2: System.Console.Write("\\"); break;
                case 3: System.Console.Write("|"); break;
            }
            Thread.Sleep(100);
            if (System.Console.CursorLeft > 0)
                System.Console.SetCursorPosition(System.Console.CursorLeft - 1, System.Console.CursorTop);
        }

        public static void PlayLoadingAnimation()
        {
            _run = true;
            Task.Run(() =>
            {
                while (_run)
                {
                    AnimateOnce();
                }
            });

        }

        public static void StopLoadingAnimation()
        {
            _run = false;
            ClearLine();
        }

        public static void ClearLine()
        {
            var currentLineCursor = System.Console.CursorTop;
            System.Console.SetCursorPosition(0, System.Console.CursorTop);
            System.Console.Write(new string(' ', System.Console.WindowWidth));
            System.Console.SetCursorPosition(0, currentLineCursor);
        }

        public static void WriteError(string value)
        {
            System.Console.ForegroundColor = ConsoleColor.Red;
            System.Console.WriteLine($"ERROR: {value}");
            System.Console.ResetColor();
        }

        public static void WriteSuccess(string value)
        {
            System.Console.ForegroundColor = ConsoleColor.Green;
            System.Console.WriteLine(value);
            System.Console.ResetColor();
        }

        public static string ReadPassword()
        {
            var password = string.Empty;
            ConsoleKeyInfo key;
            do
            {
                key = System.Console.ReadKey(true);
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar;
                    System.Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                    {
                        password = password.Substring(0, (password.Length - 1));
                        System.Console.Write("\b \b");
                    }
                    if (key.Key == ConsoleKey.Enter)
                        break;
                }
            } while (key.Key != ConsoleKey.Enter);

            return password;
        }
    }
}