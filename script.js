// Source - https://stackoverflow.com/a/57060809
// Posted by Peter
// Retrieved 2026-08-31, License - CC BY-SA 4.0

var file = document.getElementById('docpicker')
var viewer = document.getElementById('dataviewer')
file.addEventListener('change', importFile);

var json
var pNoGre
var problemeStudents = []

function importFile(evt) {
  var f = evt.target.files[0];

  if (f) {
    var r = new FileReader();
    r.onload = e => {
      var contents = processExcel(e.target.result);
      console.log(contents);

      json = contents[Object.keys(contents)[0]]

      gen_groupe();
    }
    r.readAsBinaryString(f);
  } else {
    console.log("Failed to load file");
  }
}

function processExcel(data) {
  var workbook = XLSX.read(data, {
    type: 'binary'
  });

  var firstSheet = workbook.SheetNames[0];
  var data = to_json(workbook);
  return data
};

function getPNoGre(json) {
  // Calcule le taux de non grenoblois

  let p = 0

  for (e of json) {
    if (isNoGre(e)) {
      p += 1 / json.length
    }
  }

  return p
}

function isNoGre(s) {
  try {
    return (s["N-1"] != "Université Grenoble Alpes - 38400 - Saint-Martin-d'Hères" ||
      s["Nationalité"] != "France")
  } catch (e) { return false }
}

function findBinome(s) {
  for (e of json) {
    if (e["NOM"] == s["Binome DEB"]) { return e; }
  }
}

class Groupe {
  constructor(size = null) {
    console.log(size)
    this.size = size
    this.students = []
  }

  gen() {
    let toBeIngore = []
    while (json.length > 3 & (this.students.length < this.size || this.size == null)) {

      let f = json.find((e) =>
        ((pNoGre >= getPNoGre(this.students) & isNoGre(e)) ||
          (pNoGre <= getPNoGre(this.students) & !isNoGre(e))) &
        toBeIngore.findIndex((a) => JSON.stringify(e) == JSON.stringify(a)) == -1
      )

      if (typeof f !== 'undefined') {
        try {
          this.addForBinome(f);
        } catch (e) { toBeIngore.push(f); console.error(e) }
      } else break
    }

    console.log("-----------", this.students.length, this.size)

    try { return XLSX.utils.json_to_sheet(this.students) } catch (e) { return null }
  }

  addForBinome(s) {
    if ("Num Binome" in s) {
      let bin = json.filter((e) => e["Num Binome"] == s["Num Binome"])
      if (bin.length != 2) { throw new Error("Binome non trouver") }
      bin.forEach((e) => this.add(e))
    } /* else if (!("Binome DEB" in s)) {
      let bin = json.filter((e) => !("Binome DEB" in e || "Num Binome" in e)).slice(0, 2)
      if (bin.length < 2) { throw new Error("Personne pour mettre avec cette personne") }
      bin.forEach((e) => this.add(e)) 
      } */ else throw new Error("Pas de binome")
  }

  add(s) {
    this.students.push(s)
    json = json.filter((e) => JSON.stringify(e) != JSON.stringify(s))
  }

  toString() {
    return JSON.stringify(this.students, 1, 1)
  }
}

function to_json(workbook) {
  var result = {};
  workbook.SheetNames.forEach(function (sheetName) {
    var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 0
    });

    if (roa.length) result[sheetName] = roa;
  });
  return result;
};

function gen_groupe() {

  // Elève les etudiant en Césure
  json = json.filter((e) => e == null || e["Binome DEB"] != "Césure")

  pNoGre = getPNoGre(json)

  console.log(pNoGre)

  console.log(json)

  var wb = XLSX.utils.book_new();

  gList = []
  for (e of document.getElementById("groupes").getElementsByTagName("input")) {
    gList.push(new Groupe(parseInt(e.value)))
  }

  console.log(gList)

  console.log(json)

  for (let e of json) {
    for (let i in gList) {
      if (!("Binome DEB" in e)) continue
      else if (e["Binome DEB"].split("/").slice(-1)[0].replace(" ", "").toLowerCase() == "groupe" + (parseInt(i) + 1)) {
        console.log("demande un groupe", e)
        gList[i].add(e)
      }
    }
  }

    // trouve les binome
  let bID = 1
  for (let e of json) {
    for (let s of json) {
      if (typeof e["Binome DEB"] !== 'undefined') {
        // & e["Binome DEB"].split(" ").includes(s["NOM"])

        if (e["Binome DEB"].split(" ").includes(s["NOM"])) {
          console.log("tezteznj")
          s["Num Binome"] = e["Num Binome"] = bID++;
          break;
        }
      }
    }
  }

  while (json.filter((e) => "Num Binome" in e).length > 0) {
    for (let g of gList) {
      try { g.addForBinome(json.filter((e) => "Num Binome" in e)[0]) }
      catch { console.log("sat"); break }
    }
  }

  console.log("--2")

  let testIndex = 0

  while (json.length > 0 & testIndex < 200) {
    let pNoGreLs = []
    testIndex++
    let sGList = gList.filter(function (e) {console.log(e.size, e.students.length); return e.size > e.students.length})

    console.log(sGList)

    if (sGList.length == 0) break

    for (let g of sGList) {
      pNoGreLs.push(getPNoGre(g.students) - pNoGre);
    }

    if (isNoGre(json[0])) gIndex = pNoGreLs.indexOf(Math.min(...pNoGreLs))
    else gIndex = pNoGreLs.indexOf(Math.max(...pNoGreLs))

    console.log(gIndex)

    sGList[gIndex].add(json[0]);
  }



  console.log("--3")

  for (i in gList) {
    try {
      let x = XLSX.utils.json_to_sheet(gList[i].students)
      XLSX.utils.book_append_sheet(wb, x, "groupe" + (parseInt(i) + 1));
    } catch { }
  }

  let wgp = XLSX.utils.json_to_sheet(problemeStudents.concat(json));
  XLSX.utils.book_append_sheet(wb, wgp, "groupeProbleme");

  XLSX.writeFile(wb, "groupes.xlsx");
}