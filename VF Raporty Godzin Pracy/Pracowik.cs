using System.Collections.Generic;

namespace VF_Raporty_Godzin_Pracy
{
    public class Pracowik
    {
        public string Imie;
        public string Nazwisko;
        public List<Dzien> Dni;
        public int StartIndex;
        public int KoniecIndex;

        public Pracowik()
        {
            Dni = new List<Dzien>();
        }

    }
}