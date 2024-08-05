using Datio.GenerowanieProtokolowJakoZalacznik;
using Soneta.Tools;
using System.Collections;
using Soneta.Business;
using Soneta.Business.UI;
using Soneta.Core.ServiceInvoker;
using Soneta.Business.Db;
using Soneta.Core;
using Soneta.Handel;
using Soneta.Types;
using Soneta.Support.Support;
using Soneta.Support;
using System.Reflection;

[assembly: AssemblyFileVersion("1.2.6.6")]
[assembly: AssemblyVersion("1.2.6.6")]
[assembly: Worker(typeof(GenerujZalacznik), typeof(DokHandlowe))]

namespace Datio.GenerowanieProtokolowJakoZalacznik
{
    public class GenerujZalacznik
    {
        //Podniesienie okna parametrów wydruku i przekazanie do kontekstu
        [Context]
        public Context Cx { get; set; }

        [Context]
        public Session Session { get; set; }

        [Context]
        public ParametryEksportu ParametryEksportu { get; set; }

        ArrayList Logi { get; set; }

        [Context]
        public DokumentHandlowy[] DokumentyHandlowe { get; set; }

        [Action("Import Załączników", Priority = 0, Mode = ActionMode.SingleSession, Target = ActionTarget.ToolbarWithText)]

        public object Wykonaj()
        {
            int i = 0;
            int itemzbledem = 0;
            Logi = new ArrayList();
            using (var t = DokumentyHandlowe[0].Session.Logout(true))
            {
                foreach (var item in DokumentyHandlowe)
                {
                    try
                    {
                        //Mają się drukować wszystkie załączniki
                        //               SprawdzPoprawnosc(item);
                        if (ParametryEksportu.EksportujProtokoly)
                            GenerujProtokol(item);
                        if (ParametryEksportu.EksportujFaktury)
                            GenerujFaktureSprzedazy(item);

                        var relacjaUmowa = item.NadrzedneRelacje.FirstOrDefault();
                        if (ParametryEksportu.EksportujProtokoly)
                        {
                            var pozycjeDokHandlowegoLimitHosts = relacjaUmowa?.Pozycje.FirstOrDefault()?.Nadrzedna?.UmowaLimit?.Hosts.ToList();
                            if (pozycjeDokHandlowegoLimitHosts is null)
                            {
                                itemzbledem++;
                                Logi.Add($"{Date.Now}: {item.NumerPelnyZapisany} ({item.Kontrahent.Nazwa}): Brak ustalonego limitu do pozycji na umowie");
                            }
                            else
                            {
                                i++;
                                Logi.Add($"{Date.Now}: {item.NumerPelnyZapisany} ({item.Kontrahent.Nazwa}): Plik zapisany poprawnie, weryfikacja poprawna");
                            }
                        }
                        else
                        {
                            i++;
                            Logi.Add($"{Date.Now}: {item.NumerPelnyZapisany} ({item.Kontrahent.Nazwa}): Plik zapisany poprawnie, weryfikacja poprawna");
                        }

                    }
                    catch (Exception ex)
                    {
                        itemzbledem++;
                        if (ex is Soneta.Business.DuplicatedRowException)
                        {
                            
                            Logi.Add($"{Date.Now}: {item.NumerPelnyZapisany} ({item.Kontrahent.Nazwa}): Załącznik już istnieje");
                        }
                        else
                        {
                            Logi.Add(ex.Message);
                        }                     
                    }
                }

                t.Commit();

                ZapisywanieLogow(String.Join("\n", Logi.OfType<String>()));
                string info = $"Zakończono dodawanie załączników, dodano {i} z {DokumentyHandlowe.Length}\n" +
                         $"Protokoły oraz logi zostały zapisane w C:\\Protokoły";
                if (itemzbledem != 0)
                {
                    info = $"Zakończono dodawanie załączników, dodano {i} z {DokumentyHandlowe.Length}, nie dodano {itemzbledem} z {DokumentyHandlowe.Length}.\n" +
                         $"Protokoły oraz logi zostały zapisane w C:\\Protokoły";
                }
                return new MessageBoxInformation()
                {
                    Text = info,
                };
            }
        }

        private void GenerujProtokol(DokumentHandlowy Dh)
        {
            IReportService rs;
            rs = Cx.Session.GetService<IReportService>();
            //
            //<- opcjonalnie parametry wydruku można na stałe ustawić w kontekście
            ///!!!
            Cx.Set(new PrnParams(Cx) { Kontrahent = Dh.Kontrahent, Okres = new FromTo(Dh.Data.FirstDayMonth(), Dh.Data.LastDayMonth()), DrukowanieZafakturowanych = true });
            string kodSDWorx = "00897";
            if (Dh.Kontrahent.Kod == kodSDWorx)
            {
                GenerujProtokolSDW(Dh, rs);
            }
            else
            {
                string nameW = Dh.Kontrahent.Features["Skrócona nazwa"] + "_" + "protokół" + "_" + Date.Today.Month + ".pdf";
                GenerateAndAttachReport(Dh, rs, nameW);
            }
        }

        public void GenerujProtokolSDW(DokumentHandlowy Dh, IReportService rs)
        {
            List<TicketDefinition> ticketsDefs = Session.GetSupport().TicketsDefs.WgDefaultTeam.ToList();
            var relacjaUmowa = Dh.NadrzedneRelacje.FirstOrDefault();
            var pozycjeDokHandlowegoLimitHosts = relacjaUmowa?.Pozycje.FirstOrDefault()?.Nadrzedna?.UmowaLimit?.Hosts.ToList();
            if (pozycjeDokHandlowegoLimitHosts != null)
            {
                foreach (PozycjaDokHandlowegoLimitHost pozycjaDokHandlowegoLimitHost in pozycjeDokHandlowegoLimitHosts)
                {
                    var matchingTickets = ticketsDefs.Where(ticketDef => ticketDef.ID == pozycjaDokHandlowegoLimitHost.Host.ID).ToList();
                    foreach (var ticketDef in matchingTickets)
                    {
                        Cx.Set(ticketDef);
                        string nameW = Dh.Kontrahent.Features["Skrócona nazwa"] + "_" + ticketDef.Symbol + "_" + "protokół" + "_" + Date.Today.Month + ".pdf";

                        GenerateAndAttachReport(Dh, rs, nameW);
                    }
                }
            }
        }

        private void GenerateAndAttachReport(DokumentHandlowy Dh, IReportService rs, string nameW)
        {
            var bm = BusinessModule.GetInstance(Cx.Session);

            var report = new ReportResult
            {
                TemplateFileSource = AspxSource.Storage,
                TemplateFileName = "XtraReports/Wzorce użytkownika/timetracks.repx",
                Format = ReportResultFormat.PDF,
                DataType = typeof(TimeTrack),
                Context = Cx,
                Sign = false,
                //opcjonalnie - podpis dokumentu certyfikatem kwalifikowanym, podniesienie okna wyboru certyfikatu
                VisibleSignature = false,
            };

            GenerateFile(rs.GenerateReport(report), "C:\\Protokoły\\" + nameW);

            //stworzenie obiektu klasy Attachment 
            Attachment zalW = new Attachment(Dh, AttachmentType.Attachments);

            bm.Attachments.AddRow(zalW);
            zalW.Name = nameW;
            using (Stream streamW = rs.GenerateReport(report))
            {
                zalW.Send = true;
                zalW.LoadFromStream(streamW);

            }
        }

        private void GenerujFaktureSprzedazy(DokumentHandlowy Dh)
        {
            string nameW = "FV_" + Dh.Kontrahent.Features["Skrócona nazwa"] + ".pdf";
            IReportService rs;
            rs = Cx.Session.GetService<IReportService>();

            //<- op
            //cjonalnie parametry wydruku można na stałe ustawić w kontekście

            Cx.Set(Dh);

            var bm = BusinessModule.GetInstance(Cx.Session);
            string templateRaport = "";

            if (Dh.Pozycje.Any(x=>x.Cena.Symbol == "EUR"))
            {
                Assembly assembly = Assembly.Load("Soneta.Handel.Reports");
                Type classType = assembly.GetType("Soneta.Handel.Reports.SprzedazSnippet+MyParametryWydruku");
                object classInstance = Activator.CreateInstance(classType, new object[] { Cx });
                PropertyInfo duplikatProperty = classType.GetProperty("Duplikat");
                duplikatProperty.SetValue(classInstance, false);
                PropertyInfo rabatProperty = classType.GetProperty("Rabat");
                rabatProperty.SetValue(classInstance, false);
                PropertyInfo wgCenProperty = classType.GetProperty("WgCen");
                wgCenProperty.SetValue(classInstance, false);
                Cx.Set(classInstance);
                templateRaport = "XtraReports/Wzorce użytkownika/kurs.repx";
            }
            else 
            {
                templateRaport = "XtraReports/Wzorce użytkownika/sprzedaz(dodatek).repx";
            }
            var report = new ReportResult
            {

                TemplateFileName = templateRaport,
                Format = ReportResultFormat.PDF,
                DataType = typeof(DokumentHandlowy),
                //TemplateType = typeof(XtraReportSerialization.Sprzedaz),
                Context = Cx,
                Sign = false,
                //opcjonalnie - podpis dokumentu certyfikatem kwalifikowanym, podniesienie okna wyboru certyfikatu
                VisibleSignature = false,
            };

            GenerateFile(rs.GenerateReport(report), "C:\\Protokoły\\" + nameW);

            //stworzenie obiektu klasy Attachment 
            Attachment zalW = new Attachment(Dh, AttachmentType.Attachments);

            bm.Attachments.AddRow(zalW);
            zalW.Name = nameW;
            using (Stream streamW = rs.GenerateReport(report))
            {
                zalW.Send = true;
                zalW.LoadFromStream(streamW);

            }
        }
    

        private void GenerateFile(Stream stream, string path)
        {
            var fileInfo = new FileInfo(path);

            if (!fileInfo.Directory.Exists)
                fileInfo.Directory.Create();
            using (var fh = File.Create(path))
            {
                stream.Seek(0L, SeekOrigin.Begin);
                CoreTools.StreamCopy(stream, fh);
                fh.Flush();
            }

        }
        /*     [Action("Czy fakturowane", Priority = 0, Mode = ActionMode.SingleSession, Target = ActionTarget.ToolbarWithText)]
             public void ZaznaczFakturowane()
             {
                 var timeTrack = Cx.Session.GetCore().TimeTracks;
                 using (var t = Cx.Session.Logout(true))
                 {
                     foreach (TimeTrack time in timeTrack)
                         time.Features["CzyFakturować?"] = true;
                     t.Commit();
                 }
             }*/

        public decimal SumowanieMinutNaProtokole(DokumentHandlowy dok)
        {
            var Okres = new FromTo(dok.Data.FirstDayMonth(), dok.Data.LastDayMonth());

            var czasWMin = CoreModule.GetInstance(Cx.Session).TimeTracks.WgHost
                .Where(d => ((Soneta.Support.Support.Ticket)d.GetHost(Cx.Session)).Client == dok.Kontrahent && Okres.Contains(d.Data) && (bool)d.Features["CzyFakturować?"] == true)
                .Sum(k => k.CzasWykonania.TotalMinutes);
            var godz = czasWMin / 60;
            var min = czasWMin % 60;

            return czasWMin;
        }
        

        public bool SprawdzPoprawnosc(DokumentHandlowy dokument)
        {
            var SumaMinutNaProtokole = SumowanieMinutNaProtokole(dokument);
            var SumaMinutNaFakturze = PrzeliczNaMinuty((decimal)dokument.Pozycje.Where(x => x.Kod.Contains("ZL")).Sum(p => p.Ilosc.Value));
            Logi.Add($"{Date.Now}: {dokument.NumerPelnyZapisany} ({dokument.Kontrahent.Nazwa}): Minuty na fakturze: {SumaMinutNaFakturze}");
            Logi.Add($"{Date.Now}: {dokument.NumerPelnyZapisany} ({dokument.Kontrahent.Nazwa}): Minuty na Protokole: {SumaMinutNaProtokole}");

            if (SumaMinutNaFakturze != SumaMinutNaProtokole)

                throw new Exception($"{Date.Now}: {dokument.NumerPelnyZapisany} ({dokument.Kontrahent.Nazwa}): Plik niezapisany , weryfikacja niepoprawna");
            else
                return true;
        }

        public int PrzeliczNaMinuty(decimal ilosc)
        {

            return (int)ilosc * 60;
        }

        public static void ZapisywanieLogow(string input)
        {
            File.WriteAllText($"C:\\Protokoły\\Logi_{Date.Today.Month}.txt", input);
        }
        public static bool IsVisibleWykonaj(Context context)
        {
            var dokument = (DokumentHandlowy[])context[typeof(DokumentHandlowy[])];

            if (dokument is null)
                return false;

            return dokument.Any(x=>x.Definicja.Symbol == "FV");

        }

    }

}

