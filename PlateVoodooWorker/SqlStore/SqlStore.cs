using System;

namespace PlateVoodooWorker.SqlStore
{
    public static class SqlStore
    {
        public static string AselsanSql(string transactionId, string companyId)
        {
            string SQL = @"SELECT TRANSACTIONID,LPR_REAR,LPR_REARCONF,LPR_FRONT,LPR_FRONTCONF,
                         LPR_FINAL,LPR_FINALCONF,APPROVEDPLATENUMBER,APPROVEDCLASS,TABLETYPE
                         FROM {0}.SAP_TRANSACTIONS_VIEW
                         WHERE TRANSACTIONID = '{1}'";

            return string.Format(SQL, companyId == "212" ? "KMOAVRUPA" : "KMOASYA", transactionId);
        }
    }
}