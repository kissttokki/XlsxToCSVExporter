using System.Collections;
using System.Collections.Generic;
using DataTable.Enum;

namespace ThreeKm
{
    public class $CLASSNAME : ITableData<$INDEX_TYPE>
    {
$BODY
        public $INDEX_TYPE GetKey() => $INDEX_KEY;
    }

    public class $CLASSNAMEWrapper : DataWrapper<$INDEX_TYPE, $CLASSNAME>
    {        

$BODY

        public override $CLASSNAME ToData()
        {
$CLASSNAME data = new $CLASSNAME();
$WARPPERBODY

            return data;
        }
    }


    public partial class Table$CLASSNAME : AutoBaseTableLoader<$CLASSNAMEWrapper,$INDEX_TYPE, $CLASSNAME>
    {
        public override string Path => "$CLASSNAME";
    }

}

