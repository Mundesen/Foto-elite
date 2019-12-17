using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Compare.Models
{
    class ExcelHeaders
    {
        public enum Headers
        {
            SkoleNavn, SkoleKort, Bbpakke, Aarsbillede, Elevnavn, Bestilling, Email, LandekodeForMobil, MobilForaeldre,
            HvadVilDuBestille1, HvadVilDuBestille2, HvadVilDuBestille3, HvadVilDuBestille4, Billedepakke285kr, Billedepakke285krFarveEllerSortHvid,
            Billedepakke185kr, Billedepakke185krFarveEllerSortHvid, AarsbilledeNavn, AarsbilledePris, AarsbilledeAntal, VaelgerEkstraBilledepakke,
            EkstraBilledepakke, EkstraBilledepakkeFarveEllerSortHvid, Total, OprettetAfBrugerId, IntastningsId, DatoForIndtastning, KildeUrl, TransaktionsId,
            Betalingsbeloeb, Betalingsdato, Betalingsstatus, IndlaegsId
        }

        public static Headers[] AllHeaders = 
        {
            Headers.SkoleNavn,Headers.SkoleKort,Headers.Bbpakke,Headers.Aarsbillede,Headers.Elevnavn,Headers.Bestilling,Headers.Email,Headers.LandekodeForMobil,Headers.MobilForaeldre,
           Headers.HvadVilDuBestille1,Headers.HvadVilDuBestille2,Headers.HvadVilDuBestille3,Headers.HvadVilDuBestille4,Headers.Billedepakke285kr,Headers.Billedepakke285krFarveEllerSortHvid,
           Headers.Billedepakke185kr,Headers.Billedepakke185krFarveEllerSortHvid,Headers.AarsbilledeNavn,Headers.AarsbilledePris,Headers.AarsbilledeAntal,Headers.VaelgerEkstraBilledepakke,
           Headers.EkstraBilledepakke,Headers.EkstraBilledepakkeFarveEllerSortHvid,Headers.Total,Headers.OprettetAfBrugerId,Headers.IntastningsId,Headers.DatoForIndtastning,Headers.KildeUrl,Headers.TransaktionsId,
           Headers.Betalingsbeloeb,Headers.Betalingsdato,Headers.Betalingsstatus,Headers.IndlaegsId
        };
    }
}
