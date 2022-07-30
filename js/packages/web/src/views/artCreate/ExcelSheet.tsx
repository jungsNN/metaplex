import React from "react";
import XLSX from "xlsx";

/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* Notes:
   - usage: `ReactDOM.render( <SheetJSApp />, document.getElementById('app') );`
   - xlsx.full.min.js is loaded in the head of the HTML page
   - this script should be referenced with type="text/babel"
   - babel.js in-browser transpiler should be loaded before this script
*/
const ExcelSheet = () => {
    const [data, setData] = React.useState<Array<any>>([]);
    const [cols, setCols] = React.useState<Array<any>>([]);
    const [sheets, setSheets] = React.useState<any>();
    React.useEffect(() => {
      console.log('sheets', sheets)
    }, [sheets])
  function handleFile(file: any) {
    
    /* Boilerplate to set up FileReader */
    const reader = new FileReader();
    const rABS = !!reader.readAsBinaryString;

    
    reader.onload = e => {
        if (!e.target) {
            return;
        }
      /* Parse data */
      const bstr = e.target.result;

      const wb = XLSX.read(bstr, { type: "buffer" });
      // setSheets(wb)
    //   /* Get first worksheet */
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
    //   /* Convert array of arrays */
      const data = XLSX.utils.sheet_to_json(ws, {
        header: 1
      });
    //   /* Update state */
      setData(data ?? [])
      setCols(makeCols(ws["!ref"]) ?? [])
    
    };
    // if (rABS) reader.readAsBinaryString(file);
    // else reader.readAsArrayBuffer(file);
    reader.readAsArrayBuffer(file);
  }
  function exportFile() {
    /* convert state to workbook */
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
    /* generate XLSX file and send to client */
    XLSX.writeFile(wb, "sheetjs.xlsx");
  }
  return (
    <DragDropFile handleFile={handleFile}>
      <div className="row">
        <div className="col-xs-12">
          <DataInput handleFile={handleFile} />
        </div>
      </div>
      <div className="row">
        <div className="col-xs-12">
          <button
            disabled={data.length === 0}
            className="btn btn-success"
            onClick={exportFile}
          >
            Export
          </button>
        </div>
      </div>
      <div className="row">
        <div className="col-xs-12">
          <OutTable data={data} cols={cols} />
        </div>
      </div>
      <div className="row">
        <div className="col-xs-12">
          <pre>{JSON.stringify(data, null, 1)}</pre>
        </div>
      </div>
    </DragDropFile>
  );
}

// if (typeof module !== "undefined") module.exports = SheetJSApp;

/* -------------------------------------------------------------------------- */

/*
  Simple HTML5 file drag-and-drop wrapper
  usage: <DragDropFile handleFile={handleFile}>...</DragDropFile>
    handleFile(file:File):void;
*/

// @ts-ignore
const DragDropFile = ({handleFile=(_e: File) => {}, children}) => {
 
  function suppress(evt: any) {
    evt.stopPropagation();
    evt.preventDefault();
  }
  function onDrop(evt: any) {
    evt.stopPropagation();
    evt.preventDefault();
    const files = evt.dataTransfer.files;
    if (files && files[0]) handleFile(files[0]);
  }
  return (
    <div
      onDrop={onDrop}
      onDragEnter={suppress}
      onDragOver={suppress}
    >
      {children}
    </div>
  );
}

/*
  Simple HTML5 file input wrapper
  usage: <DataInput handleFile={callback} />
    handleFile(file:File):void;
*/
// @ts-ignore
const DataInput = ({handleFile = (_e: File) => {}, }) => {
  function handleChange(e: any) {
    const files = e.target.files;
    if (files && files[0]) handleFile(files[0]);
  }
  return (
    <form className="form-inline">
      <div className="form-group">
        <label htmlFor="file">Spreadsheet</label>
        <input
          type="file"
          className="form-control"
          id="file"
          accept={SheetJSFT}
          onChange={handleChange}
        />
      </div>
    </form>
  );
}

/*
  Simple HTML Table
  usage: <OutTable data={data} cols={cols} />
    data:Array<Array<any> >;
    cols:Array<{name:string, key:number|string}>;
*/
const OutTable = ({cols, data}: {
    cols: Array<ExcelCol>;
    data: Array<any>
}) =>  {
    return (
        <div className="table-responsive">
          <table className="table table-striped">
            <thead>
              <tr>
              {/* @ts-ignore */}
                {cols.map((c: ExcelCol) => (
                  <th key={c.key}>{c.name}</th>
                ))}
              </tr>
            </thead>
            <tbody>
                {/* @ts-ignore */}
              {data.map((r: { [x: string]: boolean | React.ReactChild | React.ReactFragment | React.ReactPortal | null | undefined; }, i: React.Key | null | undefined) => (
                <tr key={i}>
                  {cols.map((c: { key: React.Key; }) => (
                    <td key={c.key}>{r[c.key]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
}

/* list of supported file types */
const SheetJSFT = [
  "xlsx",
  "xlsb",
  "xlsm",
  "xls",
  "xml",
  "csv",
  "txt",
  "ods",
  "fods",
  "uos",
  "sylk",
  "dif",
  "dbf",
  "prn",
  "qpw",
  "123",
  "wb*",
  "wq*",
  "html",
  "htm"
]
  .map(function(x) {
    return "." + x;
  })
  .join(",");


/* generate an array of column objects */
const makeCols = (refstr: string | undefined) => {
    if (typeof refstr === 'undefined') return;
  let o: Array<ExcelCol> = [],
    C = XLSX.utils.decode_range(refstr).e.c + 1;
  for (var i = 0; i < C; ++i) o.push({ name: XLSX.utils.encode_col(i), key: i });
  return o;
};

interface ExcelCol {
    name: string;
    key: number;
}

export default ExcelSheet;