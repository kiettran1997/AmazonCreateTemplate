using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AmazonCreateTemplateFile
{
    public static class ThreadSafeRandom
    {
        private static string _stringBase = "1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM";

        [ThreadStatic]
        private static Random _local;

        public static Random ThisThreadsRandom => _local ?? (_local = new Random(unchecked(Environment.TickCount * 31 + Thread.CurrentThread.ManagedThreadId)));

        public static string RandomString(int length)
        {
            var builder = new StringBuilder();
            for (int i = 0; i < length; i++)
            {
                var index = ThisThreadsRandom.Next(0, _stringBase.Length);
                builder.Append(_stringBase[index]);
            }

            return builder.ToString(0, length);
        }
    }
}
