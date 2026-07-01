using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Fo76ini.Controls
{
    public partial class UserControlHero : UserControl
    {
        // https://www.steamgriddb.com/game/5067850
        public static String FallbackHeroURL = "https://cdn.cloudflare.steamstatic.com/steam/apps/1151340/library_hero.jpg";
        public static float HeroAspectRatio = 1920f / 620f;

        public UserControlHero()
        {
            InitializeComponent();
        }

        private void UserControlHero_Load(object sender, EventArgs e)
        {
            // Fetch dynamically in background:
            Task.Run(() =>
            {
                string targetUrl = FallbackHeroURL;
                try
                {
                    using (WebClient wc = new WebClient())
                    {
                        // Add common user agent headers to bypass basic scraping blocks
                        wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.0.0 Safari/537.36");
                        string html = wc.DownloadString("https://fallout.bethesda.net/en/seasons");
                        
                        // Find all matching Cloudinary URLs in the page source
                        MatchCollection matches = Regex.Matches(html, @"https://res\.cloudinary\.com/[^""')\s]+");
                        
                        // We want to find a desktop/banner image (typically containing "Desktop", "Standard_Bnet", or "Season" and containing high-res keywords like 1200x or 1920x)
                        // and specifically filter out thumbnail parameters (like w_240, w_300, etc.)
                        string bestMatch = null;
                        foreach (Match m in matches)
                        {
                            string url = m.Value;
                            
                            // Skip small scale/transformed versions (e.g. w_240, c_lfill)
                            if (url.Contains("/w_") || url.Contains("c_lfill") || url.Contains("w_300") || url.Contains("w_240"))
                                continue;
                            
                            if (url.Contains("Desktop") || url.Contains("1200x") || url.Contains("1920x") || url.Contains("Header"))
                            {
                                bestMatch = url;
                                break;
                            }
                        }
                        
                        // Fallback to first high-res season match if no specific desktop hero found
                        if (bestMatch == null)
                        {
                            foreach (Match m in matches)
                            {
                                string url = m.Value;
                                if (url.Contains("Season") && !url.Contains("/w_") && !url.Contains("c_lfill"))
                                {
                                    bestMatch = url;
                                    break;
                                }
                            }
                        }

                        if (bestMatch != null)
                        {
                            targetUrl = bestMatch;
                        }
                    }
                }
                catch (Exception)
                {
                    // Fallback to default Steam image on any error/offline
                }

                // Load image:
                this.Invoke((MethodInvoker)delegate
                {
                    this.pictureBoxHero.LoadCompleted += (s, ev) => {
                        pictureBoxHero_Resize(this, EventArgs.Empty);
                    };
                    long timestamp = DateTimeOffset.Now.ToUnixTimeSeconds();
                    this.pictureBoxHero.LoadAsync(targetUrl + (targetUrl.Contains("?") ? "&" : "?") + "t=" + timestamp.ToString());
                });
            });
        }

        private void pictureBoxHero_Resize(object sender, EventArgs e)
        {
            // If the image is loaded, we can crop/zoom it to cover the entire control size.
            // This replicates "object-fit: cover" behavior:
            if (this.pictureBoxHero.Image != null)
            {
                float imgAspect = (float)this.pictureBoxHero.Image.Width / this.pictureBoxHero.Image.Height;
                float controlAspect = (float)this.Width / this.Height;

                if (controlAspect > imgAspect)
                {
                    // Control is wider than image aspect ratio: fit width, crop height
                    this.pictureBoxHero.Width = this.Width;
                    this.pictureBoxHero.Height = (int)(this.Width / imgAspect);
                    this.pictureBoxHero.Left = 0;
                    this.pictureBoxHero.Top = (this.Height - this.pictureBoxHero.Height) / 2;
                }
                else
                {
                    // Control is taller than image aspect ratio: fit height, crop width
                    this.pictureBoxHero.Height = this.Height;
                    this.pictureBoxHero.Width = (int)(this.Height * imgAspect);
                    this.pictureBoxHero.Top = 0;
                    this.pictureBoxHero.Left = (this.Width - this.pictureBoxHero.Width) / 2;
                }
            }
            else
            {
                // Fallback basic scaling if image isn't loaded yet
                this.pictureBoxHero.Width = this.Width;
                this.pictureBoxHero.Height = (int)(Width / HeroAspectRatio) + 5;
                this.pictureBoxHero.Left = 0;
                this.pictureBoxHero.Top = (this.Height - this.pictureBoxHero.Height) / 2;
            }
        }

        // https://stackoverflow.com/a/37764157
    }
}

