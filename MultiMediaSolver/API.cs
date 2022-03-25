using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MultiMediaSolver
{
    public static class API
    {
        public static string taskFile = Directory.GetCurrentDirectory() + "\\data\\tasks.txt";
        public static string activitiesFile = Directory.GetCurrentDirectory() + "\\data\\activities.txt";
        public static string contextsFile = Directory.GetCurrentDirectory() + "\\data\\contexts.txt";
        public static string getRandomTasks()
        {
            string outstring = "";
            string[] lines = File.ReadAllLines(taskFile);

            List<int> used = new();

            for(int i = 0; i <= 3;)
            {
                int now = new Random().Next(0, lines.Length);
                if (used.Count.Equals(0))
                {
                    used.Add(now);
                    outstring += "• " + lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }else if (!used.Contains(now))
                {
                    outstring += "• " +  lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }
            }

            return outstring;

        }

        public static string getRandomActivities()
        {
            string outstring = "";
            string[] lines = File.ReadAllLines(activitiesFile);

            List<int> used = new();

            for (int i = 0; i <= 3;)
            {
                int now = new Random().Next(0, lines.Length);
                if (used.Count.Equals(0))
                {
                    used.Add(now);
                    outstring += "• " + lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }
                else if (!used.Contains(now))
                {
                    outstring += "• " + lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }
            }

            return outstring;

        }

        public static string getRandomContexts()
        {
            string outstring = "";
            string[] lines = File.ReadAllLines(contextsFile);

            List<int> used = new();

            for (int i = 0; i <= 3;)
            {
                int now = new Random().Next(0, lines.Length);
                if (used.Count.Equals(0))
                {
                    used.Add(now);
                    outstring += "• " + lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }
                else if (!used.Contains(now))
                {
                    outstring += "• " + lines[now] + Environment.NewLine + Environment.NewLine;
                    i++;
                }
            }

            return outstring;

        }

        public static string getImageUrl(string topic)
        {
            string html = GetHtmlCode(topic);
            List<string> urls = GetUrls(html);

            var rnd = new Random();

            int randomUrl = rnd.Next(0, urls.Count - 1);

            string luckyUrl = urls[randomUrl];

            return luckyUrl;

        }

        private static string GetHtmlCode(string topic)
        {

            Task<string> result = Task.Run(() => {
                var rnd = new Random();


                string url = "https://www.google.com/search?q=" + topic + "&tbm=isch";
                string data = "";

                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko";

                var response = (HttpWebResponse)request.GetResponse();

                using (Stream dataStream = response.GetResponseStream())
                {
                    if (dataStream == null)
                        return "";
                    using (var sr = new StreamReader(dataStream))
                    {
                        data = sr.ReadToEnd();
                    }
                }
                return data;
            });

            return result.Result;
        }

        private static string GetHtml(string url)
        {
            Task<string> result = Task.Run(() => {
                var rnd = new Random();

                string data = "";

                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko";

                HttpWebResponse response = null;

                try
                {
                    response = (HttpWebResponse)request.GetResponse();
                    using (Stream dataStream = response.GetResponseStream())
                    {
                        if (dataStream == null)
                            return "";
                        using (var sr = new StreamReader(dataStream))
                        {
                            data = sr.ReadToEnd();
                        }
                    }

                }
                catch(Exception e) { }

                return data;
            });

            return result.Result;
        }

        private static List<string> GetUrls(string html)
        {
            var urls = new List<string>();

            MatchCollection matches = Regex.Matches(html, "<a.+?url=(.+?)[;'].*?>", RegexOptions.IgnoreCase);

            int made = 0;
            foreach (Match match in matches)
            {
                if(made <= 5)
                {
                    if (match.Groups[1].Value.Contains("https") || match.Groups[1].Value.Contains("http"))
                    {
                        string html2 = GetHtml(match.Groups[1].Value);

                        MatchCollection matches2 = Regex.Matches(html, "<img.+?src=[\"'](.+?)[\"'].*?>", RegexOptions.IgnoreCase);

                        foreach (Match match2 in matches2)
                        {
                            if (match2.Groups[1].Value.Contains("https") || match2.Groups[1].Value.Contains("http"))
                            {
                                if (urls.Count <= 5)
                                {
                                    if (!urls.Contains(match2.Groups[1].Value))
                                    {
                                        urls.Add(match2.Groups[1].Value);
                                    }
                                }
                            }

                        }
                        made++;
                    }
                }
            }

            return urls;
        }

        public static byte[] GetImage(string url)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            var response = (HttpWebResponse)request.GetResponse();

            using (Stream dataStream = response.GetResponseStream())
            {
                if (dataStream == null)
                    return null;
                using (var sr = new BinaryReader(dataStream))
                {
                    byte[] bytes = sr.ReadBytes(100000000);

                    return bytes;
                }
            }

            return null;
        }

    }
}

