import React from "react";
import XLSX from "xlsx";

export default class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      data: [],
      cols: []
    };
    this.fileInput = React.createRef();
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
      const dataToExport = this.validateData(data);
      if (typeof dataToExport !== "undefined") {
        let ref = "A1:D1";
        for (let i = 0; i < dataToExport.length; i++) {
          if (i === 0) {
            continue;
          }

          ref = `A1:D${i + 1}`;
        }
        this.setState({ data: dataToExport, cols: make_cols(ref) });
      }
    };
    if (rABS) reader.readAsBinaryString(file);
    else reader.readAsArrayBuffer(file);
  }

  exportFile() {
    for (let i = 0; i < this.state.data.length; i++) {
      if (i === 0) {
        continue;
      }

      const data = [this.state.data[0]];
      data.push(this.state.data[i]);

      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Invoice");
      /* generate XLSX file and send to client */
      XLSX.writeFile(wb, `invoice-${i}.xlsx`);
    }
  }

  validateData = data => {
    if (data.length === 0) {
      console.log("Required timesheet data is missing");
      return;
    }

    const dataHeaders = data[0];
    if (dataHeaders.length < 6) {
      console.log("Some required data are missing");
      return;
    }

    const requiredHeaders = [
      "Employee ID",
      "Billable Rate",
      "Project",
      "Date",
      "Start Time",
      "End Time"
    ];

    const errors = [];
    const diffHeaders = requiredHeaders.filter(x => !dataHeaders.includes(x));

    if (diffHeaders.length > 0) {
      console.log(`The headers ${diffHeaders} are not allowed!`);
      return;
    }

    const timeRegex = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;

    const dataToExport = [
      ["Employee ID", "Number of Hours", "Unit Price", "Cost"]
    ];

    for (let i = 0; i < data.length; i++) {
      if (i === 0) {
        continue;
      }
      const element = data[i];
      let date;
      try {
        date = Date.parse(element[3]);
      } catch (error) {
        console.log(error);
      }

      if (date === null) {
        errors.push(`Invalid date on row ${i + 1}`);
        continue;
      }

      const unitPrice = element[1];
      if (isNaN(unitPrice)) {
        errors.push(`Billable Rate on row ${i + 1} is not a number`);
        continue;
      }

      let startTime = element[4];
      if (!timeRegex.test(startTime)) {
        errors.push(`Start time on row ${i + 1} is invalid`);
        continue;
      }
      let endTime = element[5];
      if (!timeRegex.test(endTime)) {
        errors.push(`End time on row ${i + 1} is invalid`);
        continue;
      }

      startTime = new Date("01/01/2007 " + startTime);
      endTime = new Date("01/01/2007 " + endTime);

      let difference = endTime - startTime;
      difference = difference / 60 / 60 / 1000;

      dataToExport.push([
        element[0],
        difference.toFixed(2),
        unitPrice,
        (unitPrice * difference.toFixed(2)).toFixed(2)
      ]);
    }

    return dataToExport;
  };

  render() {
    return (
      <div className="container mt-4">
        <div className="row">
          <div className="offset-md-3 col-md-6">
            <DragDropFile handleFile={this.handleFile}>
              <div className="row mt-4">
                <div className="col-xs-12">
                  <DataInput handleFile={this.handleFile} />
                </div>
              </div>
              <div className="row mt-4">
                <p>You timesheet headers should be "Employee ID", "Billable Rate", "Project", "Date", "Start Time", and "End Time"</p>
              </div>
              <div className="row mt-2">
                <div className="col-xs-12">
                  <button
                    disabled={!this.state.data.length}
                    className="btn btn-success btn-block"
                    onClick={this.exportFile}
                  >
                    Export
                  </button>
                </div>
              </div>
              <div className="row mt-4">
                <div className="col-xs-12">
                  <OutTable data={this.state.data} cols={this.state.cols} />
                </div>
              </div>
            </DragDropFile>
          </div>
        </div>
      </div>
    );
  }
}

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
          <label htmlFor="file"></label>
          <input
            type="file"
            className="form-control"
            id="file"
            accept={SheetJSFT}
            onChange={this.handleChange}
          />
        </div>
      </form>
    );
  }
}

class OutTable extends React.Component {
  render() {
    return (
      <div className="table-responsive">
        <table className="table table-striped">
          <thead>
            <tr>
              {this.props.cols.map(c => (
                <th key={c.key}>{c.name}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {this.props.data.map((r, i) => (
              <tr key={i}>
                {this.props.cols.map(c => (
                  <td key={c.key}>{r[c.key]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}

/* list of supported file types */
const SheetJSFT = ["xlsx", "xlsb", "xlsm", "xls", "csv"]
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
