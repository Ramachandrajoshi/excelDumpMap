var sourceData = undefined;
var sampleData = undefined;
var mapping = undefined;
var onSourceChange = e => {
  var files = e.target.files,
    f = files[0];
  var reader = new FileReader();
  reader.onload = e => {
    sourceData = e.target.result;
    document.getElementById("sample").disabled = false;
  };
  reader.readAsBinaryString(f);
};
var onSampleDataChange = e => {
  var files = e.target.files,
    f = files[0];
  var reader = new FileReader();
  reader.onload = e => {
    sampleData = e.target.result;
  };
  reader.readAsBinaryString(f);
};

var onMappingChange = e => {
  var files = e.target.files,
    f = files[0];
  var reader = new FileReader();
  reader.onload = e => {
    mapping = JSON.parse(e.target.result.toUpperCase());
  };
  reader.readAsText(f);
};

var exportNow = () => {
  if (sourceData && sampleData && mapping) {
    var start = document.getElementById("startSheet").value || 1;
    var end = document.getElementById("endSheet").value || -1;
    mapData(start, end);
  } else {
      console.log(mapping);
    alert("provide all nesessory data");
  }
};

var mapData = (s, e) => {
  var shourceBook = XLSX.read(sourceData, {
    type: "binary",
    cellStyles: true,
    cellNF: true,
    cellDates: true
  });
  var sampleBook = XLSX.read(sampleData, {
    type: "binary",
    cellStyles: true,
    cellNF: true,
    cellDates: true
  });
  var currentSheetIndex = s-1;
  var endSheetIndex = e>0?e:sampleBook.SheetNames.length;
  var sourceSheet = shourceBook.Sheets[shourceBook.SheetNames[0]];
  while(currentSheetIndex < endSheetIndex){
      Object.keys(mapping).forEach(key=>{
        var sheet = sampleBook.Sheets[sampleBook.SheetNames[currentSheetIndex]];
        sheet[mapping[key]].v = sourceSheet[key+(currentSheetIndex+1)].v;
        sheet[mapping[key]].t = sourceSheet[key+(currentSheetIndex+1)].t;
        sheet[mapping[key]].w = sourceSheet[key+(currentSheetIndex+1)].w;
        sheet[mapping[key]].z = sourceSheet[key+(currentSheetIndex+1)].z;
      });
      currentSheetIndex++;
  }
  exportBook(sampleBook);
};
var exportBook = (workBook)=>{
    var wopts = { bookType: "xlsx", bookSST: false, type: "binary" };

    var wbout = XLSX.write(workBook, wopts);
  
    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    }
    saveAs(new Blob([s2ab(wbout)], { type: "" }), "test.xlsx");
}