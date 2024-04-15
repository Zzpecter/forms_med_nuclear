using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL;
using System.Data;

namespace BLL
{
    public class Hema
    {
        ManejoDatos m = new ManejoDatos();
        public DataTable Leer(string directorio)
        {
            return m.ReadExcelContent(directorio);
        }

        public string Escribir(DataTable dtDatos, string nombre, string doctor, DateTime fecha)
        {
            return m.WriteExcelHema(dtDatos, nombre, doctor, fecha);
        }

        public void PrintExcel(string filePath)
        {
            m.PrintExcel(filePath, false);
        }
    }
}
