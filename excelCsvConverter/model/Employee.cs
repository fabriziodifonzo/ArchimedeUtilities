using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FileHelpers;

namespace excelCsvConverter.model
{
  public class Employee
    {
        public string Dipartimento { get; }
        public string Direzione { get; }
        public string Ufficio { get; }
        public string AreaGiuridica { get; }
        public string Cognome { get; }
        public string Nome { get; }
        public string CodiceFiscale { get; }
        public string IndirizzoMail { get; }
        public string TipoMobilita { get; }
        public string DataCessazione { get; }
        public string CausaleCessazione { get; }

        public static Employee Of(string dipartimento, string direzione, string ufficio, string areaGiuridica, string cognome, string nome, string codiceFiscale, string indirizzoMail, string tipoMobilita, string dataCessazione, string causaleCessazione)
        {
            return new Employee(dipartimento, direzione, ufficio, areaGiuridica, cognome,  nome,  codiceFiscale,  indirizzoMail, tipoMobilita, dataCessazione, causaleCessazione);
        }

        public override string ToString()
        {
            return $"[ Dipartimento: <{Dipartimento}> Direzione: <{Direzione}> Ufficio: <{Ufficio}> Area Giuridica: <{AreaGiuridica}> Cognome: <{Cognome}> Nome: <{Nome}> Codice Fiscale: <{CodiceFiscale}> Indirizzo Mail: <{IndirizzoMail}> Tipo Mobilita: <{TipoMobilita}> Data Cessazione: <{DataCessazione}> Causale Cessazione: <{CausaleCessazione}> ]";
        }

        private Employee(string dipartimento, string direzione, string ufficio, string areaGiuridica, string cognome, string nome, string codiceFiscale, string indirizzoMail, string tipoMobilita, string dataCessazione, string causaleCessazione)
        {
            Dipartimento = dipartimento;
            Direzione = direzione;
            Ufficio = ufficio;
            AreaGiuridica = areaGiuridica;
            Cognome = cognome;
            Nome = nome;
            CodiceFiscale = codiceFiscale;
            IndirizzoMail = indirizzoMail;
            TipoMobilita = tipoMobilita;
            DataCessazione = dataCessazione;
            CausaleCessazione = causaleCessazione;
        }
    }
}
