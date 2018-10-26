using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateZConsol
{
    internal class UploadDump
    {
        public DateTime dateTime;

        public string devision;

        public int worker;

        public int prostoy;

        public int po;

        public int so;

        public int kom;

        public int bol;

        public int uvo;

        public int work8;

        public int work10;

        public int setUSR;

        public int setUSMK;

        public int setUIS;

        public int setUSS;

        public int setEMU;

        public int getUSR;

        public int getUSMK;

        public int getUIS;

        public int getUSS;

        public int getEMU;

        public UploadDump(DateTime dateTime, string davisionName)
        {
            this.dateTime = dateTime;
            devision = davisionName;
            worker = 0;
            prostoy = 0;
            po = 0;
            so = 0;
            kom = 0;
            bol = 0;
            uvo = 0;
            work8 = 0;
            work10 = 0;
            setUSR = 0;
            setUSMK = 0;
            setUIS = 0;
            setUSS = 0;
            setEMU = 0;
            getUSR = 0;
            getUSMK = 0;
            getUIS = 0;
            getUSS = 0;
            getEMU = 0;
        }
    }
}
