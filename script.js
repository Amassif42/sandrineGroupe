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
  return (s["N-1"] != "Université Grenoble Alpes - 38400 - Saint-Martin-d'Hères" ||
    s["Nationalité"] != "France")
}

class Groupe {
  constructor(size = null) {
    this.size = size
    this.students = []
  }

  gen() {

    let index = 0

    while (json.length > 0 & (this.students.length < this.size || this.size == null)) {
      index++
      for (let e of json) {
        if (pNoGre - getPNoGre(this.students) > 0.1) {
          if (isNoGre(e)) { this.add(e); break }
        } else if (pNoGre - getPNoGre(this.students) < -0.1) {
          if (!isNoGre(e)) { this.add(e); break }
        } else {
          this.add(e); break
        }
      }
    }
    console.log("-----------", this.students.length, this.size)

    return XLSX.utils.json_to_sheet(this.students);
  }

  add(s, recursive = true) {
    this.students.push(s)

    let deleteFalg = false
    json = json.filter((e) => JSON.stringify(e) != JSON.stringify(s))

    if (!recursive) return

    if (s["Binome DEB"] == "Validé" || s["Binome DEB"] == "pas DEB en présentiel") return

    if("Binome DEB" in s) {
      for(e of json) {
        if(e["NOM"] == s["Binome DEB"]) {this.add(e, false); return;}
      }
    } else {
      for(e of json) {
        if(!("Binome DEB" in e)) {this.add(e, false); return;}
      }      
    }

/*
    for (e of json) {
      if ("Binome DEB" in s & "NOM" in e) if (s["Binome DEB"].toUpperCase() == e["NOM"].toUpperCase()) {
        this.add(e, false); return;
      } else if (!("Binome DEB" in s) & !("Binome DEB" in e)) { this.add(e, false); return; }
    } */

    // throw new Error("zhgeiurhezbrh")

    console.log("aaaaaaaaaaaaaaaaaaaaaaaa", s, json)
    problemeStudents.push(s)
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

  for(e of json.filter((e) => "Binome DEB" in e)) {
    for(s of json) {if (s["NOM"] == e["Binome DEB"]) {s["Binome DEB"] = e["NOM"]}}
  }

  pNoGre = getPNoGre(json)

  console.log(pNoGre)

  console.log(json)

  var wb = XLSX.utils.book_new();

  let g1 = new Groupe(6)
  let wg1 = g1.gen()

  let g2 = new Groupe(Math.floor(json.length / 2))
  let wg2 = g2.gen()

  let g3 = new Groupe(null)
  let wg3 = g3.gen()

  XLSX.utils.book_append_sheet(wb, wg1, "groupe1");

  XLSX.utils.book_append_sheet(wb, wg2, "groupe2");

  XLSX.utils.book_append_sheet(wb, wg3, "groupe3");

  let wgp = XLSX.utils.json_to_sheet(problemeStudents);
  XLSX.utils.book_append_sheet(wb, wgp, "groupeProbleme");

  XLSX.writeFile(wb, "SheetJSExportAOO.xlsx");
}