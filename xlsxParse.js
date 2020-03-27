var xlsx = require('xlsx');

module.exports = () => {
  const getSubsection = (ws0, from, to) => {
    const range = to - from;
    var value = [];
    for (let i = 0; i <= range; i++) {
      value[i] = ws0[`A${i + from}`] ? ws0[`A${i + from}`].v : undefined;
    }
    return value;
  };
  if (xlsx.readFile('./public/uploads/data.xlsx')) {
    var wb = xlsx.readFile('./public/uploads/data.xlsx');
    const sheetNames = wb.SheetNames;

    var ws0 = wb.Sheets[`${sheetNames[2]}`];

    var file = [];
    //SECTION VALUE
    var sectionValue = [];
    sectionValue[0] = ws0['A2'] ? ws0['A2'].v : undefined;
    sectionValue[1] = ws0['A12'] ? ws0['A12'].v : undefined;
    sectionValue[2] = ws0['A20'] ? ws0['A20'].v : undefined;
    sectionValue[3] = ws0['A57'] ? ws0['A57'].v : undefined;
    sectionValue[4] = ws0['A68'] ? ws0['A68'].v : undefined;
    sectionValue[5] = ws0['A104'] ? ws0['A104'].v : undefined;
    sectionValue[6] = ws0['A111'] ? ws0['A111'].v : undefined;
    sectionValue[7] = ws0['A122'] ? ws0['A122'].v : undefined;

    //SIBSECTION VALUE
    var subsectionValue = [];
    subsectionValue[0] = getSubsection(ws0, 3, 10);
    subsectionValue[1] = getSubsection(ws0, 13, 18);
    subsectionValue[2] = getSubsection(ws0, 21, 55);
    subsectionValue[3] = getSubsection(ws0, 58, 66);
    subsectionValue[4] = getSubsection(ws0, 69, 102);
    subsectionValue[5] = getSubsection(ws0, 105, 109);
    subsectionValue[6] = getSubsection(ws0, 112, 120);
    subsectionValue[7] = getSubsection(ws0, 123, 143);

    //merged array of subsection
    var subMerge = [];
    subMerge = subMerge.concat(sectionValue[0]);
    subMerge = subMerge.concat(subsectionValue[0]);
    subMerge = subMerge.concat(sectionValue[1]);
    subMerge = subMerge.concat(subsectionValue[1]);
    subMerge = subMerge.concat(sectionValue[2]);
    subMerge = subMerge.concat(subsectionValue[2]);
    subMerge = subMerge.concat(sectionValue[3]);
    subMerge = subMerge.concat(subsectionValue[3]);
    subMerge = subMerge.concat(sectionValue[4]);
    subMerge = subMerge.concat(subsectionValue[4]);
    subMerge = subMerge.concat(sectionValue[5]);
    subMerge = subMerge.concat(subsectionValue[5]);
    subMerge = subMerge.concat(sectionValue[6]);
    subMerge = subMerge.concat(subsectionValue[6]);
    subMerge = subMerge.concat(sectionValue[7]);
    subMerge = subMerge.concat(subsectionValue[7]);

    for (let i = 0; i < subMerge.length; i++) {
      file[i] = {
        Requirement: subMerge[i],
        Response: ''
      };
    }
  }

  return file;
};
