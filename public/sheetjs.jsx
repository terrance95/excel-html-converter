/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* Notes:
   - usage: `ReactDOM.render( <SheetJSApp />, document.getElementById('app') );`
   - xlsx.full.min.js is loaded in the head of the HTML page
   - this script should be referenced with type="text/babel"
   - babel.js in-browser transpiler should be loaded before this script
*/
class SheetJSApp extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      data: [] /* Array of Arrays e.g. [["a","b"],[1,2]] */,
      cols: [] /* Array of column objects e.g. { name: "C", K: 2 } */
    };
    this.handleFile = this.handleFile.bind(this);
    this.exportFile = this.exportFile.bind(this);
  }
  handleFile(file /*:File*/) {
    /* Boilerplate to set up FileReader */
    const reader = new FileReader();
    const rABS = !!reader.readAsBinaryString;
    reader.onload = e => {
      /* Parse data */
      const bstr = e.target.result;
      const wb = XLSX.read(bstr, { type: rABS ? "binary" : "array" });
      /* Get first worksheet */
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      /* Convert array of arrays */
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      /* Update state */
      this.setState({ data: data, cols: make_cols(ws["!ref"]) });
    };
    if (rABS) reader.readAsBinaryString(file);
    else reader.readAsArrayBuffer(file);
  }
  exportFile() {
    /* convert state to workbook */
    const ws = XLSX.utils.aoa_to_sheet(this.state.data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
    /* generate XLSX file and send to client */
    XLSX.writeFile(wb, "sheetjs.xlsx");
  }
  render() {
    return (
      <DragDropFile handleFile={this.handleFile}>
        <div className="row">
          <div className="col-xs-12">
            <DataInput handleFile={this.handleFile} />
          </div>
        </div>

        <div className="row">
          <br />
          <div className="col-xs-12">
            <OutHTML data={this.state.data} cols={this.state.cols} />
          </div>
        </div>
      </DragDropFile>
    );
  }
}

if (typeof module !== "undefined") module.exports = SheetJSApp;

/* -------------------------------------------------------------------------- */

/*
  Simple HTML5 file drag-and-drop wrapper
  usage: <DragDropFile handleFile={handleFile}>...</DragDropFile>
    handleFile(file:File):void;
*/
class DragDropFile extends React.Component {
  constructor(props) {
    super(props);
    this.onDrop = this.onDrop.bind(this);
  }
  suppress(evt) {
    evt.stopPropagation();
    evt.preventDefault();
  }
  onDrop(evt) {
    evt.stopPropagation();
    evt.preventDefault();
    const files = evt.dataTransfer.files;
    if (files && files[0]) this.props.handleFile(files[0]);
  }
  render() {
    return (
      <div
        onDrop={this.onDrop}
        onDragEnter={this.suppress}
        onDragOver={this.suppress}
      >
        {this.props.children}
      </div>
    );
  }
}

/*
  Simple HTML5 file input wrapper
  usage: <DataInput handleFile={callback} />
    handleFile(file:File):void;
*/
class DataInput extends React.Component {
  constructor(props) {
    super(props);
    this.handleChange = this.handleChange.bind(this);
  }
  handleChange(e) {
    const files = e.target.files;
    if (files && files[0]) this.props.handleFile(files[0]);
  }
  render() {
    return (
      <form className="form-inline">
        <div className="form-group">
          <label style={{ color: "#fff" }} htmlFor="file">
            Spreadsheet
          </label>
          <br />
          <input
            type="file"
            className="form-control"
            id="file"
            accept={SheetJSFT}
            onChange={this.handleChange}
          />
          <br />
        </div>
      </form>
    );
  }
}

/*
  Simple HTML Table
  usage: <OutTable data={data} cols={cols} />
    data:Array<Array<any> >;
    cols:Array<{name:string, key:number|string}>;
*/
class OutHTML extends React.Component {
  constructor(props) {
    super(props);
  }
  render() {
    return (
      <pre className="line-numbers language-markup">
        <code>
          &lt;div class="slide-bar-box"&gt;
          {this.props.data.map(item => (
            <div>
              <br />
              &lt;div class="media"&gt;
              <br />
              &nbsp;&nbsp;&lt;div class="media-left"&gt;
              <br />
              &nbsp;&nbsp;&nbsp;&lt;img src="/static/images/article-thumbnails/
              {item[3]}.jpg" alt="" class="media-object"&gt;
              <br />
              &nbsp;&nbsp;&lt;/div&gt;
              <br />
              <br />
              &nbsp;&lt;div class="media-body"&gt;
              <br />
              &nbsp;&nbsp;&nbsp;&lt;h3 class="title"&gt;&lt;span&gt;
              {item[0] + ": "}
              &lt;/span&gt;&lt;a href="{item[5]}" class="eng" rel="external"
              title="{"英語の原文を読む："}
              {item[1]}"&gt;{item[2]}&lt;/a&gt;&lt;/h3&gt;
              <br />
              &nbsp;&nbsp;&lt;p class="eng-title"&gt;
              {item[1]}
              &nbsp;&nbsp;&lt;/p&gt;
              <br />
              &nbsp;&nbsp;&lt;p class="doi"&gt;{"doi: 10.1038/"}
              {item[3]}&lt;/p&gt;
              <br />
              &nbsp;&nbsp;&lt;p class="standfirst"&gt;{item[4]}&lt;/p&gt;
              <br />
              &nbsp;&lt;/div&gt;
              <br />
              &lt;/div&gt;
              <br />
              <br />
              <br />
            </div>
          ))}
          <br /> &lt;/div&gt;
        </code>
      </pre>
    );
  }
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
const make_cols = refstr => {
  let o = [],
    C = XLSX.utils.decode_range(refstr).e.c + 1;
  for (var i = 0; i < C; ++i) o[i] = { name: XLSX.utils.encode_col(i), key: i };
  return o;
};
