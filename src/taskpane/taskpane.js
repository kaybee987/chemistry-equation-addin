Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    init();
  }
});

let reactants = [];
let products = [];
let lastFocusedInput = null;
let selectedFormulaId = null;
let activeView = "equation-builder-view";

// Token types: n=normal, i=italic, b=subscript, p=superscript, f=fraction {num:[tokens], den:[tokens]}
const FORMULA_DATA = [
  // === Gas Laws ===
  { id:"ideal-gas", name:"Ideal Gas Law", category:"Gas Laws",
    tokens:[{t:"PV",y:"i"},{t:" = ",y:"n"},{t:"nRT",y:"i"}],
    desc:"Relates pressure, volume, amount and temperature of an ideal gas.",
    vars:[["P","pressure"],["V","volume"],["n","moles"],["R","gas constant (8.314 J/mol·K)"],["T","temperature (K)"]]},
  { id:"combined-gas", name:"Combined Gas Law", category:"Gas Laws",
    tokens:[{y:"f",num:[{t:"P",y:"i"},{t:"1",y:"b"},{t:"V",y:"i"},{t:"1",y:"b"}],den:[{t:"T",y:"i"},{t:"1",y:"b"}]},{t:" = ",y:"n"},{y:"f",num:[{t:"P",y:"i"},{t:"2",y:"b"},{t:"V",y:"i"},{t:"2",y:"b"}],den:[{t:"T",y:"i"},{t:"2",y:"b"}]}],
    desc:"Combines Boyle's, Charles's and Gay-Lussac's laws for a fixed amount of gas.",
    vars:[["P","pressure"],["V","volume"],["T","temperature (K)"]]},
  { id:"dalton", name:"Dalton's Law of Partial Pressures", category:"Gas Laws",
    tokens:[{t:"P",y:"i"},{t:"total",y:"b"},{t:" = P",y:"n"},{t:"1",y:"b"},{t:" + P",y:"n"},{t:"2",y:"b"},{t:" + ... + P",y:"n"},{t:"n",y:"b"}],
    desc:"Total pressure is the sum of partial pressures of all gases in a mixture.",
    vars:[["Pₜₒₜₐₗ","total pressure"],["Pᵢ","partial pressure of gas i"]]},
  { id:"graham", name:"Graham's Law of Effusion", category:"Gas Laws",
    tokens:[{y:"f",num:[{t:"r",y:"i"},{t:"1",y:"b"}],den:[{t:"r",y:"i"},{t:"2",y:"b"}]},{t:" = √",y:"n"},{y:"f",num:[{t:"M",y:"i"},{t:"2",y:"b"}],den:[{t:"M",y:"i"},{t:"1",y:"b"}]}],
    desc:"Rate of effusion is inversely proportional to the square root of molar mass.",
    vars:[["r","rate of effusion"],["M","molar mass"]]},

  // === Acid-Base ===
  { id:"henderson", name:"Henderson-Hasselbalch Equation", category:"Acid-Base",
    tokens:[{t:"pH = p",y:"n"},{t:"K",y:"i"},{t:"a",y:"b"},{t:" + log",y:"n"},{y:"f",num:[{t:"[A",y:"n"},{t:"−",y:"p"},{t:"]",y:"n"}],den:[{t:"[HA]",y:"n"}]}],
    desc:"Relates pH of a buffer to pKa and ratio of conjugate base to acid.",
    vars:[["pH","negative log of [H⁺]"],["pKa","acid dissociation constant"],["[A⁻]","conjugate base concentration"],["[HA]","weak acid concentration"]]},
  { id:"ka-kb", name:"Ka × Kb Relationship", category:"Acid-Base",
    tokens:[{t:"K",y:"i"},{t:"a",y:"b"},{t:" × ",y:"n"},{t:"K",y:"i"},{t:"b",y:"b"},{t:" = ",y:"n"},{t:"K",y:"i"},{t:"w",y:"b"},{t:" = 1.0 × 10",y:"n"},{t:"−14",y:"p"}],
    desc:"Product of Ka and Kb for a conjugate acid-base pair equals Kw.",
    vars:[["Ka","acid dissociation constant"],["Kb","base dissociation constant"],["Kw","water autoionization constant"]]},
  { id:"ph-poh", name:"pH + pOH Relationship", category:"Acid-Base",
    tokens:[{t:"pH + pOH = 14",y:"n"}],
    desc:"At 25°C, the sum of pH and pOH equals 14.",
    vars:[["pH","−log[H⁺]"],["pOH","−log[OH⁻]"]]},
  { id:"ph-def", name:"pH Definition", category:"Acid-Base",
    tokens:[{t:"pH = −log[H",y:"n"},{t:"+",y:"p"},{t:"]",y:"n"}],
    desc:"pH is the negative logarithm of hydrogen ion concentration.",
    vars:[["pH","measure of acidity"],["[H⁺]","hydrogen ion concentration (mol/L)"]]},

  // === Thermodynamics ===
  { id:"gibbs", name:"Gibbs Free Energy", category:"Thermodynamics",
    tokens:[{t:"ΔG",y:"n"},{t:"°",y:"p"},{t:" = ΔH",y:"n"},{t:"°",y:"p"},{t:" − ",y:"n"},{t:"T",y:"i"},{t:"ΔS",y:"n"},{t:"°",y:"p"}],
    desc:"Relates free energy change to enthalpy, temperature and entropy.",
    vars:[["ΔG°","standard free energy change"],["ΔH°","standard enthalpy change"],["T","temperature (K)"],["ΔS°","standard entropy change"]]},
  { id:"gibbs-nonstandard", name:"Gibbs Free Energy (non-standard)", category:"Thermodynamics",
    tokens:[{t:"ΔG = ΔG",y:"n"},{t:"°",y:"p"},{t:" + ",y:"n"},{t:"RT",y:"i"},{t:" ln ",y:"n"},{t:"Q",y:"i"}],
    desc:"Free energy under non-standard conditions using reaction quotient.",
    vars:[["ΔG","free energy change"],["ΔG°","standard free energy change"],["R","gas constant"],["T","temperature (K)"],["Q","reaction quotient"]]},
  { id:"hess", name:"Hess's Law", category:"Thermodynamics",
    tokens:[{t:"ΔH",y:"n"},{t:"rxn",y:"b"},{t:" = ΣΔH",y:"n"},{t:"f",y:"b"},{t:"°(products) − ΣΔH",y:"n"},{t:"f",y:"b"},{t:"°(reactants)",y:"n"}],
    desc:"Enthalpy change equals sum of enthalpies of formation of products minus reactants.",
    vars:[["ΔHᵣₓₙ","reaction enthalpy"],["ΔHf°","standard enthalpy of formation"]]},
  { id:"entropy", name:"Boltzmann Entropy", category:"Thermodynamics",
    tokens:[{t:"S",y:"i"},{t:" = ",y:"n"},{t:"k",y:"i"},{t:"B",y:"b"},{t:" ln ",y:"n"},{t:"W",y:"i"}],
    desc:"Entropy as a function of the number of microstates.",
    vars:[["S","entropy"],["kB","Boltzmann constant (1.38×10⁻²³ J/K)"],["W","number of microstates"]]},
  { id:"clausius", name:"Clausius-Clapeyron Equation", category:"Thermodynamics",
    tokens:[{t:"ln",y:"n"},{y:"f",num:[{t:"P",y:"i"},{t:"2",y:"b"}],den:[{t:"P",y:"i"},{t:"1",y:"b"}]},{t:" = −",y:"n"},{y:"f",num:[{t:"ΔH",y:"n"},{t:"vap",y:"b"}],den:[{t:"R",y:"i"}]},{t:"(",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"T",y:"i"},{t:"2",y:"b"}]},{t:" − ",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"T",y:"i"},{t:"1",y:"b"}]},{t:")",y:"n"}],
    desc:"Relates vapor pressure to temperature and enthalpy of vaporization.",
    vars:[["P","vapor pressure"],["ΔHᵥₐₚ","enthalpy of vaporization"],["R","gas constant"],["T","temperature (K)"]]},

  // === Kinetics ===
  { id:"arrhenius", name:"Arrhenius Equation", category:"Kinetics",
    tokens:[{t:"k",y:"i"},{t:" = ",y:"n"},{t:"A",y:"i"},{t:"e",y:"n"},{t:"−Ea/RT",y:"p"}],
    desc:"Rate constant as a function of temperature and activation energy.",
    vars:[["k","rate constant"],["A","pre-exponential factor"],["Ea","activation energy"],["R","gas constant"],["T","temperature (K)"]]},
  { id:"rate-first", name:"First-Order Rate Law", category:"Kinetics",
    tokens:[{t:"ln[A] = ln[A]",y:"n"},{t:"0",y:"b"},{t:" − ",y:"n"},{t:"kt",y:"i"}],
    desc:"Integrated rate law for first-order reactions.",
    vars:[["[A]","concentration at time t"],["[A]₀","initial concentration"],["k","rate constant"],["t","time"]]},
  { id:"half-life-first", name:"First-Order Half-Life", category:"Kinetics",
    tokens:[{t:"t",y:"i"},{t:"½",y:"b"},{t:" = ",y:"n"},{y:"f",num:[{t:"0.693",y:"n"}],den:[{t:"k",y:"i"}]}],
    desc:"Half-life for a first-order reaction is independent of concentration.",
    vars:[["t½","half-life"],["k","rate constant"]]},
  { id:"rate-second", name:"Second-Order Rate Law", category:"Kinetics",
    tokens:[{y:"f",num:[{t:"1",y:"n"}],den:[{t:"[A]",y:"n"}]},{t:" = ",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"[A]",y:"n"},{t:"0",y:"b"}]},{t:" + ",y:"n"},{t:"kt",y:"i"}],
    desc:"Integrated rate law for second-order reactions.",
    vars:[["[A]","concentration at time t"],["[A]₀","initial concentration"],["k","rate constant"],["t","time"]]},

  // === Electrochemistry ===
  { id:"nernst", name:"Nernst Equation", category:"Electrochemistry",
    tokens:[{t:"E",y:"i"},{t:" = ",y:"n"},{t:"E",y:"i"},{t:"°",y:"p"},{t:" − ",y:"n"},{y:"f",num:[{t:"RT",y:"i"}],den:[{t:"nF",y:"i"}]},{t:" ln ",y:"n"},{t:"Q",y:"i"}],
    desc:"Cell potential under non-standard conditions.",
    vars:[["E","cell potential (V)"],["E°","standard cell potential"],["R","gas constant"],["T","temperature (K)"],["n","moles of electrons"],["F","Faraday constant (96485 C/mol)"],["Q","reaction quotient"]]},
  { id:"nernst-25", name:"Nernst Equation (25°C)", category:"Electrochemistry",
    tokens:[{t:"E",y:"i"},{t:" = ",y:"n"},{t:"E",y:"i"},{t:"°",y:"p"},{t:" − ",y:"n"},{y:"f",num:[{t:"0.0592",y:"n"}],den:[{t:"n",y:"i"}]},{t:" log ",y:"n"},{t:"Q",y:"i"}],
    desc:"Simplified Nernst equation at 25°C using log base 10.",
    vars:[["E","cell potential (V)"],["E°","standard cell potential"],["n","moles of electrons"],["Q","reaction quotient"]]},
  { id:"faraday", name:"Faraday's Law of Electrolysis", category:"Electrochemistry",
    tokens:[{t:"m",y:"i"},{t:" = ",y:"n"},{y:"f",num:[{t:"MIt",y:"i"}],den:[{t:"nF",y:"i"}]}],
    desc:"Mass deposited during electrolysis.",
    vars:[["m","mass deposited"],["M","molar mass"],["I","current (A)"],["t","time (s)"],["n","electrons per ion"],["F","Faraday constant"]]},

  // === Equilibrium ===
  { id:"kp-kc", name:"Kp and Kc Relationship", category:"Equilibrium",
    tokens:[{t:"K",y:"i"},{t:"p",y:"b"},{t:" = ",y:"n"},{t:"K",y:"i"},{t:"c",y:"b"},{t:"(",y:"n"},{t:"RT",y:"i"},{t:")",y:"n"},{t:"Δn",y:"p"}],
    desc:"Relates equilibrium constants in terms of pressure and concentration.",
    vars:[["Kp","equilibrium constant (pressure)"],["Kc","equilibrium constant (concentration)"],["R","gas constant"],["T","temperature (K)"],["Δn","change in moles of gas"]]},
  { id:"vant-hoff", name:"Van't Hoff Equation", category:"Equilibrium",
    tokens:[{t:"ln",y:"n"},{y:"f",num:[{t:"K",y:"i"},{t:"2",y:"b"}],den:[{t:"K",y:"i"},{t:"1",y:"b"}]},{t:" = ",y:"n"},{y:"f",num:[{t:"ΔH",y:"n"},{t:"°",y:"p"}],den:[{t:"R",y:"i"}]},{t:"(",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"T",y:"i"},{t:"1",y:"b"}]},{t:" − ",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"T",y:"i"},{t:"2",y:"b"}]},{t:")",y:"n"}],
    desc:"Temperature dependence of equilibrium constant.",
    vars:[["K","equilibrium constant"],["ΔH°","standard enthalpy change"],["R","gas constant"],["T","temperature (K)"]]},

  // === Quantum / Spectroscopy ===
  { id:"rydberg", name:"Rydberg Equation", category:"Quantum",
    tokens:[{y:"f",num:[{t:"1",y:"n"}],den:[{t:"λ",y:"i"}]},{t:" = ",y:"n"},{t:"R",y:"i"},{t:"H",y:"b"},{t:"(",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"n",y:"i"},{t:"1",y:"b"},{t:"²",y:"p"}]},{t:" − ",y:"n"},{y:"f",num:[{t:"1",y:"n"}],den:[{t:"n",y:"i"},{t:"2",y:"b"},{t:"²",y:"p"}]},{t:")",y:"n"}],
    desc:"Wavelength of light emitted by hydrogen atom electron transitions.",
    vars:[["λ","wavelength"],["Rₕ","Rydberg constant (1.097×10⁷ m⁻¹)"],["n","principal quantum numbers"]]},
  { id:"de-broglie", name:"de Broglie Wavelength", category:"Quantum",
    tokens:[{t:"λ = ",y:"n"},{y:"f",num:[{t:"h",y:"i"}],den:[{t:"mv",y:"i"}]}],
    desc:"Wavelength associated with a moving particle.",
    vars:[["λ","wavelength"],["h","Planck's constant (6.626×10⁻³⁴ J·s)"],["m","mass"],["v","velocity"]]},
  { id:"planck", name:"Planck's Equation", category:"Quantum",
    tokens:[{t:"E",y:"i"},{t:" = ",y:"n"},{t:"hν",y:"i"}],
    desc:"Energy of a photon is proportional to its frequency.",
    vars:[["E","energy"],["h","Planck's constant"],["ν","frequency (Hz)"]]},
  { id:"heisenberg", name:"Heisenberg Uncertainty Principle", category:"Quantum",
    tokens:[{t:"Δ",y:"n"},{t:"x",y:"i"},{t:" · Δ",y:"n"},{t:"p",y:"i"},{t:" ≥ ",y:"n"},{y:"f",num:[{t:"h",y:"i"}],den:[{t:"4π",y:"n"}]}],
    desc:"Fundamental limit on precision of position and momentum measurements.",
    vars:[["Δx","uncertainty in position"],["Δp","uncertainty in momentum"],["h","Planck's constant"]]},

  // === Solutions ===
  { id:"raoult", name:"Raoult's Law", category:"Solutions",
    tokens:[{t:"P",y:"i"},{t:"A",y:"b"},{t:" = ",y:"n"},{t:"χ",y:"i"},{t:"A",y:"b"},{t:"P",y:"i"},{t:"A",y:"b"},{t:"°",y:"p"}],
    desc:"Vapor pressure of a solution component equals mole fraction times pure vapor pressure.",
    vars:[["Pₐ","partial vapor pressure"],["χₐ","mole fraction"],["Pₐ°","vapor pressure of pure component"]]},
  { id:"dilution", name:"Dilution Equation", category:"Solutions",
    tokens:[{t:"M",y:"i"},{t:"1",y:"b"},{t:"V",y:"i"},{t:"1",y:"b"},{t:" = ",y:"n"},{t:"M",y:"i"},{t:"2",y:"b"},{t:"V",y:"i"},{t:"2",y:"b"}],
    desc:"Moles of solute remain constant during dilution.",
    vars:[["M","molarity"],["V","volume"]]},
  { id:"molality", name:"Molality Definition", category:"Solutions",
    tokens:[{t:"m",y:"i"},{t:" = ",y:"n"},{y:"f",num:[{t:"mol solute",y:"n"}],den:[{t:"kg solvent",y:"n"}]}],
    desc:"Molality is moles of solute per kilogram of solvent.",
    vars:[["m","molality (mol/kg)"]]},
  { id:"boiling-elevation", name:"Boiling Point Elevation", category:"Solutions",
    tokens:[{t:"ΔT",y:"n"},{t:"b",y:"b"},{t:" = ",y:"n"},{t:"i",y:"i"},{t:"K",y:"i"},{t:"b",y:"b"},{t:"m",y:"i"}],
    desc:"Boiling point increase due to dissolved solute.",
    vars:[["ΔTb","boiling point elevation"],["i","van't Hoff factor"],["Kb","ebullioscopic constant"],["m","molality"]]},
  { id:"freezing-depression", name:"Freezing Point Depression", category:"Solutions",
    tokens:[{t:"ΔT",y:"n"},{t:"f",y:"b"},{t:" = ",y:"n"},{t:"i",y:"i"},{t:"K",y:"i"},{t:"f",y:"b"},{t:"m",y:"i"}],
    desc:"Freezing point decrease due to dissolved solute.",
    vars:[["ΔTf","freezing point depression"],["i","van't Hoff factor"],["Kf","cryoscopic constant"],["m","molality"]]},
];

function init() {
  addReactant();
  addProduct();

  document.getElementById("btn-add-reactant").addEventListener("click", addReactant);
  document.getElementById("btn-add-product").addEventListener("click", addProduct);
  document.getElementById("btn-insert").addEventListener("click", handleInsert);
  document.getElementById("btn-clear").addEventListener("click", clearAll);

  document.getElementById("arrow-type").addEventListener("change", updatePreview);
  document.getElementById("arrow-head").addEventListener("change", updatePreview);
  document.getElementById("above-arrow").addEventListener("input", updatePreview);
  document.getElementById("below-arrow").addEventListener("input", updatePreview);
  document.getElementById("arrow-length").addEventListener("input", updateArrowSlider);

  document.getElementById("above-arrow").addEventListener("focus", (e) => { lastFocusedInput = e.target; });
  document.getElementById("below-arrow").addEventListener("focus", (e) => { lastFocusedInput = e.target; });

  document.getElementById("above-arrow").addEventListener("blur", (e) => {
    const corrected = autoCapitalizeFormula(e.target.value);
    if (corrected !== e.target.value) { e.target.value = corrected; updatePreview(); }
  });
  document.getElementById("below-arrow").addEventListener("blur", (e) => {
    const corrected = autoCapitalizeFormula(e.target.value);
    if (corrected !== e.target.value) { e.target.value = corrected; updatePreview(); }
  });

  setupSymbolLibrary();
  initMainTabs();
  initFormulaReference();
  updatePreview();
}

function updateArrowSlider() {
  const val = document.getElementById("arrow-length").value;
  document.getElementById("arrow-length-val").textContent = val + "px";
  updatePreview();
}

function createComponentCard(type, index) {
  const card = document.createElement("div");
  card.className = "component-card";
  card.dataset.type = type;
  card.dataset.index = index;

  card.innerHTML = `
    <button class="btn-remove" title="Remove">&times;</button>
    <div class="card-fields">
      <div>
        <label>Coeff.</label>
        <input type="text" class="coeff-input" placeholder="1" maxlength="4"/>
      </div>
      <div class="formula-field">
        <label>Formula</label>
        <input type="text" class="formula-input" placeholder="e.g. h2so4"/>
      </div>
      <div>
        <label>Charge</label>
        <input type="text" class="charge-input" placeholder="2+" maxlength="4"/>
      </div>
      <div>
        <label>State</label>
        <input type="text" class="state-input" placeholder="aq" maxlength="6"/>
      </div>
    </div>
    <div class="card-name">
      <label>Name / Number</label>
      <input type="text" class="name-input" placeholder="e.g. sulfuric acid, compound 1"/>
    </div>
  `;

  card.querySelector(".btn-remove").addEventListener("click", () => {
    removeComponent(type, card);
  });

  card.querySelectorAll("input").forEach((input) => {
    input.addEventListener("input", updatePreview);
    input.addEventListener("focus", () => {
      lastFocusedInput = input;
    });
  });

  const formulaInput = card.querySelector(".formula-input");
  formulaInput.addEventListener("blur", () => {
    const corrected = autoCapitalizeFormula(formulaInput.value);
    if (corrected !== formulaInput.value) {
      formulaInput.value = corrected;
      updatePreview();
    }
  });

  return card;
}

function addReactant() {
  const list = document.getElementById("reactants-list");
  const index = list.children.length;
  const card = createComponentCard("reactant", index);
  list.appendChild(card);
  reactants.push(card);
  updatePreview();
}

function addProduct() {
  const list = document.getElementById("products-list");
  const index = list.children.length;
  const card = createComponentCard("product", index);
  list.appendChild(card);
  products.push(card);
  updatePreview();
}

function removeComponent(type, card) {
  card.remove();
  if (type === "reactant") {
    reactants = reactants.filter((r) => r !== card);
  } else {
    products = products.filter((p) => p !== card);
  }
  updatePreview();
}

function getComponentData(card) {
  return {
    coeff: card.querySelector(".coeff-input").value.trim(),
    formula: card.querySelector(".formula-input").value.trim(),
    charge: card.querySelector(".charge-input").value.trim(),
    state: card.querySelector(".state-input").value.trim(),
    name: card.querySelector(".name-input").value.trim(),
  };
}

function formatFormula(formula) {
  return formula
    .replace(/([A-Za-z)\]])(\d+)/g, (_, prev, digits) => {
      return prev + digits
        .split("")
        .map((d) => String.fromCharCode(0x2080 + parseInt(d)))
        .join("");
    });
}

function formatCharge(charge) {
  if (!charge) return "";
  const superMap = {
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴",
    "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "+": "⁺", "-": "⁻", "−": "⁻",
  };
  return charge.split("").map((c) => superMap[c] || c).join("");
}

function formatStatePreview(state) {
  if (!state) return "";
  const clean = state.replace(/[()]/g, "").trim();
  if (!clean) return "";
  return `<sub>(${clean})</sub>`;
}

function formatStateText(state) {
  if (!state) return "";
  const clean = state.replace(/[()]/g, "").trim();
  if (!clean) return "";
  return "(" + clean + ")";
}

// Common chemical groups — matched FIRST (before individual elements)
// Ordered longest-first so longer groups match before shorter ones
const COMMON_GROUPS = [
  "COOH","COO","CO3","CO2","CO",
  "CH3","CH2","CHO","CN","CS",
  "SO4","SO3","SO2","SO",
  "NO3","NO2","NO",
  "NH4","NH3","NH2","NH",
  "OH","PO4","PO3",
  "ClO4","ClO3","ClO2","ClO",
  "CrO4","Cr2O7",
  "MnO4",
  "HCO3","HSO4","HPO4","H2PO4",
  "SiO4","SiO3","SiO2",
  "BrO3","IO3",
];

// All elements — 2-letter first, then 1-letter
const ELEMENTS_2 = [
  "He","Li","Be","Ne","Na","Mg","Al","Si","Cl","Ar","Ca","Sc","Ti",
  "Cr","Mn","Fe","Co","Ni","Cu","Zn","Ga","Ge","As","Se","Br","Kr",
  "Rb","Sr","Zr","Nb","Mo","Tc","Ru","Rh","Pd","Ag","Cd","In","Sn",
  "Sb","Te","Xe","Cs","Ba","La","Ce","Pr","Nd","Pm","Sm","Eu","Gd",
  "Tb","Dy","Ho","Er","Tm","Yb","Lu","Hf","Ta","Re","Os","Ir","Pt",
  "Au","Hg","Tl","Pb","Bi","Po","At","Rn","Fr","Ra","Ac","Th","Pa",
  "Np","Pu","Am","Cm","Bk","Cf","Es","Fm","Md","No","Lr","Rf","Db",
  "Sg","Bh","Hs","Mt","Ds","Rg","Cn","Nh","Fl","Mc","Lv","Ts","Og",
];

const ELEMENTS_1 = ["H","B","C","N","O","F","P","S","K","V","Y","I","W","U"];

const ENGLISH_WORDS = new Set([
  // Common chemistry context words
  "acid","base","salt","water","heat","cold","room","temp","temperature",
  "hot","warm","cool","boil","boiling","freeze","freezing","melt","melting",
  "dilute","diluted","concentrated","conc","dil","excess","limited",
  "catalyst","catalytic","enzyme","light","dark","pressure","high","low",
  "aqueous","liquid","solid","gas","vapor","vapour","solution","solvent",
  "solute","mixture","pure","impure","dry","wet","filter","distill",
  "reflux","stir","stirring","shake","slow","fast","rapid","gentle",
  "vigorous","overnight","hours","minutes","seconds","days",
  "add","added","remove","removed","dissolve","dissolved","mix","mixed",
  "react","reacted","produce","produced","yield","gives","forms",
  "oxidize","oxidized","reduce","reduced","neutralize","neutralized",
  "burn","burned","decompose","decomposed","evaporate","sublime",
  "precipitate","precipitated","crystallize","saturate","saturated",
  "unsaturated","supersaturated","anhydrous","hydrated","molten",
  "finely","divided","powdered","granular","concentrated","standard",
  "normal","strong","weak","polar","nonpolar","ionic","covalent",
  "organic","inorganic","metallic","alkaline","acidic","neutral",
  "exothermic","endothermic","spontaneous","reversible","irreversible",
  "equilibrium","complete","incomplete","partial","total",
  "dropwise","portion","slowly","quickly","carefully","immediately",
  "above","below","over","under","with","without","into","from",
  "the","and","then","not","for","this","that","step",
  "air","sun","sunlight","uv","infrared",
]);

function isEnglishText(input) {
  const words = input.toLowerCase().split(/[\s,;:]+/).filter(Boolean);
  if (words.length === 0) return false;
  return words.some((w) => ENGLISH_WORDS.has(w));
}

function autoCapitalizeFormula(input) {
  if (!input) return input;
  if (!document.getElementById("autoformat-enabled").checked) return input;
  if (isEnglishText(input)) return input;

  let result = "";
  let i = 0;
  const str = input;

  while (i < str.length) {
    // Skip numbers, brackets, parentheses, dots — keep as-is
    if (/[\d()\[\]·.+\-]/.test(str[i])) {
      result += str[i];
      i++;
      continue;
    }

    let matched = false;

    // 1. Try common groups first (longest match wins)
    for (const group of COMMON_GROUPS) {
      const chunk = str.substring(i, i + group.length);
      if (chunk.toLowerCase() === group.toLowerCase()) {
        // Verify next char is not a lowercase letter (would mean it's part of something else)
        const nextChar = str[i + group.length];
        const nextIsLower = nextChar && /[a-z]/.test(nextChar);
        if (!nextIsLower) {
          result += group;
          i += group.length;
          matched = true;
          break;
        }
      }
    }

    if (matched) continue;

    // 2. Try 2-letter elements
    if (i + 1 < str.length) {
      const twoChar = str.substring(i, i + 2);
      // Only match 2-letter element if second char is lowercase (or we'll force it)
      for (const el of ELEMENTS_2) {
        if (twoChar.toLowerCase() === el.toLowerCase()) {
          result += el;
          i += 2;
          matched = true;
          break;
        }
      }
    }

    if (matched) continue;

    // 3. Try 1-letter elements
    const oneChar = str[i];
    for (const el of ELEMENTS_1) {
      if (oneChar.toLowerCase() === el.toLowerCase()) {
        result += el;
        i++;
        matched = true;
        break;
      }
    }

    if (!matched) {
      result += str[i];
      i++;
    }
  }

  return result;
}

function buildPreviewText() {
  const above = document.getElementById("above-arrow").value.trim();
  const below = document.getElementById("below-arrow").value.trim();

  const reactantParts = [];
  const reactantNames = [];
  reactants.forEach((card) => {
    const data = getComponentData(card);
    if (!data.formula) return;
    let part = "";
    if (data.coeff && data.coeff !== "1") part += escapeHtml(data.coeff);
    part += escapeHtml(formatFormula(data.formula));
    if (data.charge) part += `<sup>${escapeHtml(data.charge)}</sup>`;
    if (data.state) part += formatStatePreview(data.state);
    reactantParts.push(part);
    reactantNames.push(data.name || "");
  });

  const productParts = [];
  const productNames = [];
  products.forEach((card) => {
    const data = getComponentData(card);
    if (!data.formula) return;
    let part = "";
    if (data.coeff && data.coeff !== "1") part += escapeHtml(data.coeff);
    part += escapeHtml(formatFormula(data.formula));
    if (data.charge) part += `<sup>${escapeHtml(data.charge)}</sup>`;
    if (data.state) part += formatStatePreview(data.state);
    productParts.push(part);
    productNames.push(data.name || "");
  });

  if (reactantParts.length === 0 && productParts.length === 0) {
    return null;
  }

  const reactantStr = reactantParts.join(" + ");
  const productStr = productParts.join(" + ");

  return { reactantStr, productStr, above, below, reactantNames, productNames };
}

function updatePreview() {
  const preview = document.getElementById("equation-preview");
  const data = buildPreviewText();

  if (!data) {
    preview.innerHTML = '<span class="placeholder">Preview will appear here...</span>';
    return;
  }

  const arrowType = document.getElementById("arrow-type").value;
  const headStyle = document.getElementById("arrow-head").value;
  const length = parseInt(document.getElementById("arrow-length").value);
  const above = document.getElementById("above-arrow").value.trim();
  const below = document.getElementById("below-arrow").value.trim();

  let arrowClass = "ap-arrow";
  if (arrowType === "reverse" || arrowType === "harpoon-left") arrowClass += " reverse";
  if (arrowType === "equilibrium" || arrowType === "reversible") arrowClass += " equilibrium";
  if (arrowType === "reversible") arrowClass += " wide";
  if (arrowType === "resonance") arrowClass += " resonance";
  if (arrowType === "no-reaction") arrowClass += " no-reaction";
  if (arrowType === "harpoon-right" || arrowType === "harpoon-left") arrowClass += " harpoon";
  arrowClass += ` head-${headStyle}`;

  let html = '<div class="preview-equation">';
  html += `<div class="pe-above">${above ? formatPreviewFormula(above) : '&nbsp;'}</div>`;
  html += `<div class="pe-reactants">${data.reactantStr}</div>`;
  html += `<div class="pe-arrow"><span class="${arrowClass}" style="width:${length}px"></span></div>`;
  html += `<div class="pe-products">${data.productStr}</div>`;
  html += `<div class="pe-below">${below ? formatPreviewFormula(below) : '&nbsp;'}</div>`;
  html += '</div>';

  const allNames = [...data.reactantNames, ...data.productNames].filter((n) => n);
  if (allNames.length > 0) {
    html += `<div class="preview-names">${allNames.map(escapeHtml).join("  |  ")}</div>`;
  }

  preview.innerHTML = html;
}

function formatPreviewFormula(text) {
  return escapeHtml(text).replace(/([A-Za-z)\]])(\d+)/g, (_, prev, digits) => {
    return prev + `<sub>${digits}</sub>`;
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function clearAll() {
  document.getElementById("reactants-list").innerHTML = "";
  document.getElementById("products-list").innerHTML = "";
  reactants = [];
  products = [];
  document.getElementById("above-arrow").value = "";
  document.getElementById("below-arrow").value = "";
  document.getElementById("arrow-type").selectedIndex = 0;
  addReactant();
  addProduct();
  updatePreview();
}

function setupSymbolLibrary() {
  document.querySelectorAll(".symbol-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".symbol-tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".symbol-panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(`panel-${tab.dataset.tab}`).classList.add("active");
    });
  });

  document.querySelectorAll(".sym-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const symbol = btn.dataset.symbol;
      insertSymbolAtCursor(symbol);
    });
  });
}

function insertSymbolAtCursor(symbol) {
  if (lastFocusedInput) {
    const input = lastFocusedInput;
    const start = input.selectionStart;
    const end = input.selectionEnd;
    const value = input.value;
    input.value = value.substring(0, start) + symbol + value.substring(end);
    input.selectionStart = input.selectionEnd = start + symbol.length;
    input.focus();
    updatePreview();
  }
}

async function insertEquation() {
  const data = buildPreviewText();
  if (!data) {
    showStatus("Please add at least one reactant or product.", "error");
    return;
  }

  try {
    await Word.run(async (context) => {
      const ooxml = buildOoxml(data);
      const range = context.document.getSelection();
      const inserted = range.insertOoxml(ooxml, Word.InsertLocation.replace);
      const afterRange = inserted.getRange("After");
      afterRange.select();
      await context.sync();
      showStatus("Equation inserted successfully!", "success");
      document.activeElement.blur();
    });
  } catch (error) {
    showStatus("Error inserting equation: " + error.message, "error");
  }
}

function buildOoxml(data) {
  const arrowType = document.getElementById("arrow-type").value;
  const above = document.getElementById("above-arrow").value.trim();
  const below = document.getElementById("below-arrow").value.trim();
  const arrowLengthPx = parseInt(document.getElementById("arrow-length").value);
  const lineCharCount = Math.max(5, Math.round(arrowLengthPx / 8));

  const cellBorders = `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>`;
  const rowHeight = 240;
  const trPr = `<w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="exact"/></w:trPr>`;
  const arrowWidthDxa = lineCharCount * 120;

  // Build columns: each reactant, each +, the arrow block, each product, each +
  // Structure: [R1] [+] [R2] [+] ... [arrow] ... [P1] [+] [P2]
  // Each column has: row1=above (merged for arrow), row2=formula, row3=below (merged for arrow), row4=name

  const columns = []; // { type: 'component'|'plus'|'arrow', runs: string, name: string }

  // Reactants
  const activeReactants = [];
  reactants.forEach((card) => {
    const comp = getComponentData(card);
    if (!comp.formula) return;
    activeReactants.push(comp);
  });
  activeReactants.forEach((comp, i) => {
    if (i > 0) columns.push({ type: "plus", runs: makeRun("+", false, false), name: "" });
    let runs = "";
    if (comp.coeff && comp.coeff !== "1") runs += makeRun(comp.coeff, false, false);
    runs += buildFormulaRuns(comp.formula);
    if (comp.charge) runs += makeChargeRun(comp.charge);
    if (comp.state) runs += makeRun("(" + comp.state.replace(/[()]/g, "") + ")", true, false);
    columns.push({ type: "component", runs, name: comp.name || "" });
  });

  // Arrow
  columns.push({ type: "arrow", runs: buildArrowRuns(arrowType, lineCharCount), name: "" });

  // Products
  const activeProducts = [];
  products.forEach((card) => {
    const comp = getComponentData(card);
    if (!comp.formula) return;
    activeProducts.push(comp);
  });
  activeProducts.forEach((comp, i) => {
    if (i > 0) columns.push({ type: "plus", runs: makeRun("+", false, false), name: "" });
    let runs = "";
    if (comp.coeff && comp.coeff !== "1") runs += makeRun(comp.coeff, false, false);
    runs += buildFormulaRuns(comp.formula);
    if (comp.charge) runs += makeChargeRun(comp.charge);
    if (comp.state) runs += makeRun("(" + comp.state.replace(/[()]/g, "") + ")", true, false);
    columns.push({ type: "component", runs, name: comp.name || "" });
  });

  const totalCols = columns.length;
  const arrowIdx = columns.findIndex((c) => c.type === "arrow");

  // Row 1: above-arrow text (merged across arrow column, vMerge restart for others)
  let row1 = `<w:tr>${trPr}`;
  columns.forEach((col, i) => {
    if (i === arrowIdx) {
      row1 += `<w:tc><w:tcPr><w:tcW w:w="${arrowWidthDxa}" w:type="dxa"/><w:vAlign w:val="center"/>${cellBorders}</w:tcPr>`;
      row1 += above ? makeParagraph(buildSmallFormulaRuns(above), "center") : makeParagraph("", "center");
      row1 += `</w:tc>`;
    } else {
      row1 += `<w:tc><w:tcPr><w:vMerge w:val="restart"/><w:vAlign w:val="center"/>${cellBorders}</w:tcPr>`;
      row1 += makeParagraph(col.runs, "center");
      row1 += `</w:tc>`;
    }
  });
  row1 += `</w:tr>`;

  // Row 2: arrow in arrow column, vMerge continue for others
  let row2 = `<w:tr>${trPr}`;
  columns.forEach((col, i) => {
    if (i === arrowIdx) {
      row2 += `<w:tc><w:tcPr><w:tcW w:w="${arrowWidthDxa}" w:type="dxa"/>${cellBorders}</w:tcPr>`;
      row2 += `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr>${col.runs}</w:p>`;
      row2 += `</w:tc>`;
    } else {
      row2 += `<w:tc><w:tcPr><w:vMerge/>${cellBorders}</w:tcPr><w:p/></w:tc>`;
    }
  });
  row2 += `</w:tr>`;

  // Row 3: below-arrow text in arrow column, vMerge continue for others
  let row3 = `<w:tr>${trPr}`;
  columns.forEach((col, i) => {
    if (i === arrowIdx) {
      row3 += `<w:tc><w:tcPr><w:tcW w:w="${arrowWidthDxa}" w:type="dxa"/><w:vAlign w:val="center"/>${cellBorders}</w:tcPr>`;
      row3 += below ? makeParagraph(buildSmallFormulaRuns(below), "center") : makeParagraph("", "center");
      row3 += `</w:tc>`;
    } else {
      row3 += `<w:tc><w:tcPr><w:vMerge/>${cellBorders}</w:tcPr><w:p/></w:tc>`;
    }
  });
  row3 += `</w:tr>`;

  // Row 4: names below each component
  const hasAnyName = columns.some((c) => c.name);
  let row4 = "";
  if (hasAnyName) {
    row4 = `<w:tr><w:trPr><w:trHeight w:val="200" w:hRule="atLeast"/></w:trPr>`;
    columns.forEach((col, i) => {
      row4 += `<w:tc><w:tcPr>${i === arrowIdx ? `<w:tcW w:w="${arrowWidthDxa}" w:type="dxa"/>` : ""}${cellBorders}</w:tcPr>`;
      if (col.name) {
        row4 += makeParagraph(makeRun(col.name, false, false, true, true), "center");
      } else {
        row4 += makeParagraph("", "center");
      }
      row4 += `</w:tc>`;
    });
    row4 += `</w:tr>`;
  }

  const tbl = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="nil"/>
      <w:left w:val="nil"/>
      <w:bottom w:val="nil"/>
      <w:right w:val="nil"/>
      <w:insideH w:val="nil"/>
      <w:insideV w:val="nil"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/>
      <w:left w:w="36" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/>
      <w:right w:w="36" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  ${row1}${row2}${row3}${row4}
</w:tbl>`;

  // Wrap table with paragraphs before and after to prevent table merging
  let paragraphs = "";
  paragraphs += `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
  paragraphs += tbl;
  paragraphs += `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;

  return wrapInOoxml(paragraphs);
}

function buildArrowRuns(arrowType, charCount) {
  const headStyle = document.getElementById("arrow-head").value;
  const arrowLengthPx = parseInt(document.getElementById("arrow-length").value);
  const lengthPt = Math.round(arrowLengthPx * 0.75);

  switch (arrowType) {
    case "forward":
      return buildVmlCustomArrow(lengthPt, headStyle, "right");
    case "reverse":
      return buildVmlCustomArrow(lengthPt, headStyle, "left");
    case "equilibrium":
      return buildVmlCustomEquilibrium(lengthPt, headStyle, false);
    case "reversible":
      return buildVmlCustomEquilibrium(lengthPt, headStyle, true);
    case "resonance":
      return buildVmlCustomArrow(lengthPt, headStyle, "both");
    case "no-reaction":
      return buildVmlNoReaction(lengthPt, headStyle);
    case "harpoon-right":
      return buildVmlCustomArrow(lengthPt, "harpoon", "right");
    case "harpoon-left":
      return buildVmlCustomArrow(lengthPt, "harpoon", "left");
    default:
      return buildVmlCustomArrow(lengthPt, headStyle, "right");
  }
}

function buildVmlCustomArrow(lengthPt, headStyle, direction) {
  const heightPt = 8;
  const coordW = lengthPt * 20;
  const coordH = heightPt * 20;
  const yMid = Math.round(coordH / 2);
  const headSize = Math.round(coordH * 0.45); // arrowhead height from center
  const headLen = Math.round(coordW * 0.06); // arrowhead length (% of total)
  const minHeadLen = 80;
  const hl = Math.max(headLen, minHeadLen);

  let shapes = "";

  // Main line (shaft)
  shapes += `<v:line from="0,${yMid}" to="${coordW},${yMid}" strokeweight="1pt" strokecolor="#000000"/>`;

  // Right arrowhead
  if (direction === "right" || direction === "both") {
    shapes += buildHeadShape(headStyle, coordW, yMid, headSize, hl, "right", false);
  }

  // Left arrowhead
  if (direction === "left" || direction === "both") {
    shapes += buildHeadShape(headStyle, coordW, yMid, headSize, hl, "left", false);
  }

  return `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:group style="width:${lengthPt}pt;height:${heightPt}pt" coordsize="${coordW},${coordH}">${shapes}</v:group></w:pict></w:r>`;
}

function buildHeadShape(style, coordW, yMid, headSize, headLen, dir, flipV) {
  // tip = point of arrow, base = where it connects to shaft
  const tip = dir === "right" ? coordW : 0;
  const base = dir === "right" ? coordW - headLen : headLen;
  // flipV inverts top/bottom (for equilibrium bottom arrow)
  const top = flipV ? yMid + headSize : yMid - headSize;
  const bot = flipV ? yMid - headSize : yMid + headSize;

  switch (style) {
    case "filled": {
      const points = `${tip},${yMid} ${base},${top} ${base},${bot} ${tip},${yMid}`;
      return `<v:polyline points="${points}" fillcolor="#000000" strokecolor="#000000" strokeweight="0.5pt"><v:fill type="solid"/></v:polyline>`;
    }
    case "classic": {
      const narrowSize = Math.round(headSize * 0.6);
      const longLen = Math.round(headLen * 1.4);
      const b = dir === "right" ? coordW - longLen : longLen;
      const t = flipV ? yMid + narrowSize : yMid - narrowSize;
      const bt = flipV ? yMid - narrowSize : yMid + narrowSize;
      const points = `${tip},${yMid} ${b},${t} ${b},${bt} ${tip},${yMid}`;
      return `<v:polyline points="${points}" fillcolor="#000000" strokecolor="#000000" strokeweight="0.5pt"><v:fill type="solid"/></v:polyline>`;
    }
    case "open": {
      const points = `${base},${top} ${tip},${yMid} ${base},${bot}`;
      return `<v:polyline points="${points}" filled="f" strokecolor="#000000" strokeweight="1pt"/>`;
    }
    case "barbed": {
      const notch = dir === "right" ? coordW - Math.round(headLen * 0.5) : Math.round(headLen * 0.5);
      const points = `${tip},${yMid} ${base},${top} ${notch},${yMid} ${base},${bot} ${tip},${yMid}`;
      return `<v:polyline points="${points}" fillcolor="#000000" strokecolor="#000000" strokeweight="0.5pt"><v:fill type="solid"/></v:polyline>`;
    }
    case "harpoon": {
      // Half-arrow: only top half (or bottom half if flipped)
      const halfSide = flipV ? bot : top;
      const points = `${base},${halfSide} ${tip},${yMid} ${base},${yMid}`;
      return `<v:polyline points="${points}" fillcolor="#000000" strokecolor="#000000" strokeweight="0.5pt"><v:fill type="solid"/></v:polyline>`;
    }
    default:
      return "";
  }
}

function buildVmlCustomEquilibrium(lengthPt, headStyle, wide) {
  const heightPt = wide ? 12 : 8;
  const coordW = lengthPt * 20;
  const coordH = heightPt * 20;
  const y1 = wide ? Math.round(coordH * 0.25) : Math.round(coordH * 0.33);
  const y2 = wide ? Math.round(coordH * 0.75) : Math.round(coordH * 0.67);

  // Map custom head styles to VML stroke arrow types for reliability
  const vmlMap = { filled: "block", classic: "classic", open: "open", barbed: "block", harpoon: "classic" };
  const vmlEnd = vmlMap[headStyle] || "block";

  let shapes = "";

  // Top line: arrow on right end (endarrow)
  shapes += `<v:line from="0,${y1}" to="${coordW},${y1}" strokeweight="0.75pt" strokecolor="#000000"><v:stroke endarrow="${vmlEnd}" endarrowwidth="medium" endarrowlength="medium"/></v:line>`;

  // Bottom line: arrow on left end (startarrow), vertically it's just the reverse direction
  shapes += `<v:line from="0,${y2}" to="${coordW},${y2}" strokeweight="0.75pt" strokecolor="#000000"><v:stroke startarrow="${vmlEnd}" startarrowwidth="medium" startarrowlength="medium"/></v:line>`;

  return `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:group style="width:${lengthPt}pt;height:${heightPt}pt" coordsize="${coordW},${coordH}">${shapes}</v:group></w:pict></w:r>`;
}

function buildVmlNoReaction(lengthPt, headStyle) {
  const heightPt = 8;
  const coordW = lengthPt * 20;
  const coordH = heightPt * 20;
  const yMid = Math.round(coordH / 2);
  const headSize = Math.round(coordH * 0.45);
  const headLen = Math.max(Math.round(coordW * 0.06), 80);

  let shapes = "";

  // Main shaft
  shapes += `<v:line from="0,${yMid}" to="${coordW},${yMid}" strokeweight="1pt" strokecolor="#000000"/>`;

  // Right arrowhead
  shapes += buildHeadShape(headStyle, coordW, yMid, headSize, headLen, "right", false);

  // X cross in the middle
  const cx = Math.round(coordW / 2);
  const crossSize = Math.round(coordH * 0.4);
  const crossW = Math.round(coordW * 0.03);
  shapes += `<v:line from="${cx - crossW},${yMid - crossSize}" to="${cx + crossW},${yMid + crossSize}" strokeweight="1.5pt" strokecolor="#000000"/>`;
  shapes += `<v:line from="${cx - crossW},${yMid + crossSize}" to="${cx + crossW},${yMid - crossSize}" strokeweight="1.5pt" strokecolor="#000000"/>`;

  return `<w:r><w:rPr><w:noProof/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:group style="width:${lengthPt}pt;height:${heightPt}pt" coordsize="${coordW},${coordH}">${shapes}</v:group></w:pict></w:r>`;
}

function makeArrowRun(text) {
  return `<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:spacing w:val="-20"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function makeChargeRun(text) {
  return `<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:position w:val="8"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function buildFormulaRuns(formula) {
  let runs = "";
  // Tokenize: letters, digits-after-letter/bracket (subscript), ^superscript, other
  const tokens = [];
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === "^") {
      // Superscript: collect until whitespace or end
      let sup = "";
      i++;
      while (i < formula.length && /[^\s{}]/.test(formula[i])) {
        sup += formula[i];
        i++;
      }
      tokens.push({ type: "super", text: sup });
    } else if (/\d/.test(formula[i]) && i > 0 && /[A-Za-z)\]]/.test(formula[i - 1])) {
      // Digits after a letter or closing bracket — subscript
      let digits = "";
      while (i < formula.length && /\d/.test(formula[i])) {
        digits += formula[i];
        i++;
      }
      tokens.push({ type: "sub", text: digits });
    } else {
      // Normal text — collect until we hit a subscript-eligible digit or ^
      let text = "";
      while (i < formula.length && formula[i] !== "^") {
        if (/\d/.test(formula[i]) && i > 0 && /[A-Za-z)\]]/.test(formula[i - 1])) break;
        text += formula[i];
        i++;
      }
      if (text) tokens.push({ type: "normal", text });
    }
  }

  for (const tok of tokens) {
    if (tok.type === "sub") {
      runs += makeRun(tok.text, true, false);
    } else if (tok.type === "super") {
      runs += makeRun(tok.text, false, true);
    } else {
      runs += makeRun(tok.text, false, false);
    }
  }

  return runs;
}

function buildSmallFormulaRuns(formula) {
  let runs = "";
  const tokens = [];
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === "^") {
      let sup = "";
      i++;
      while (i < formula.length && /[^\s{}]/.test(formula[i])) {
        sup += formula[i];
        i++;
      }
      tokens.push({ type: "super", text: sup });
    } else if (/\d/.test(formula[i]) && i > 0 && /[A-Za-z)\]]/.test(formula[i - 1])) {
      let digits = "";
      while (i < formula.length && /\d/.test(formula[i])) {
        digits += formula[i];
        i++;
      }
      tokens.push({ type: "sub", text: digits });
    } else {
      let text = "";
      while (i < formula.length && formula[i] !== "^") {
        if (/\d/.test(formula[i]) && i > 0 && /[A-Za-z)\]]/.test(formula[i - 1])) break;
        text += formula[i];
        i++;
      }
      if (text) tokens.push({ type: "normal", text });
    }
  }

  for (const tok of tokens) {
    if (tok.type === "sub") {
      runs += makeRun(tok.text, true, false, true, false);
    } else if (tok.type === "super") {
      runs += makeRun(tok.text, false, true, true, false);
    } else {
      runs += makeRun(tok.text, false, false, true, false);
    }
  }

  return runs;
}

function makeRun(text, subscript, superscript, small, italic) {
  let rPr = "<w:rPr>";
  rPr += '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>';

  if (small) {
    rPr += '<w:sz w:val="16"/><w:szCs w:val="16"/>';
  } else {
    rPr += '<w:sz w:val="22"/><w:szCs w:val="22"/>';
  }

  if (subscript) {
    rPr += '<w:vertAlign w:val="subscript"/>';
  }
  if (superscript) {
    rPr += '<w:vertAlign w:val="superscript"/>';
  }
  if (italic) {
    rPr += "<w:i/><w:iCs/>";
  }
  rPr += "</w:rPr>";

  const escaped = escapeXml(text);
  return `<w:r>${rPr}<w:t xml:space="preserve">${escaped}</w:t></w:r>`;
}

function makeParagraph(runs, alignment) {
  let pPr = "<w:pPr>";
  if (alignment) {
    pPr += `<w:jc w:val="${alignment}"/>`;
  }
  pPr += '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>';
  pPr += "</w:pPr>";

  return `<w:p>${pPr}${runs}</w:p>`;
}

function wrapInOoxml(bodyContent) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                  xmlns:v="urn:schemas-microsoft-com:vml"
                  xmlns:o="urn:schemas-microsoft-com:office:office">
        <w:body>
          ${bodyContent}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

function escapeXml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function showStatus(message, type) {
  let statusEl = document.querySelector(".status-msg");
  if (!statusEl) {
    statusEl = document.createElement("div");
    statusEl.className = "status-msg";
    document.querySelector(".fixed-bottom").appendChild(statusEl);
  }
  statusEl.textContent = message;
  statusEl.className = `status-msg ${type}`;
  setTimeout(() => {
    statusEl.className = "status-msg";
  }, 3000);
}

// ========== MAIN TABS ==========

function initMainTabs() {
  document.querySelectorAll(".main-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".main-tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".main-view").forEach((v) => v.classList.remove("active"));
      tab.classList.add("active");
      const target = tab.dataset.target;
      document.getElementById(target).classList.add("active");
      activeView = target;
      const btn = document.getElementById("btn-insert");
      btn.textContent = target === "formula-reference-view" ? "Insert Formula" : "Insert into Document";
    });
  });
}

function handleInsert() {
  if (activeView === "formula-reference-view") {
    insertSelectedFormula();
  } else {
    insertEquation();
  }
}

// ========== FORMULA REFERENCE ==========

function initFormulaReference() {
  const categories = [...new Set(FORMULA_DATA.map((f) => f.category))];
  const container = document.getElementById("formula-categories");
  let html = '<button class="category-chip active" data-category="all">All</button>';
  categories.forEach((cat) => {
    html += `<button class="category-chip" data-category="${cat}">${cat}</button>`;
  });
  container.innerHTML = html;

  container.querySelectorAll(".category-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      container.querySelectorAll(".category-chip").forEach((c) => c.classList.remove("active"));
      chip.classList.add("active");
      renderFormulaList();
    });
  });

  document.getElementById("formula-search-input").addEventListener("input", renderFormulaList);

  renderFormulaList();
}

function renderFormulaList() {
  const search = document.getElementById("formula-search-input").value.toLowerCase().trim();
  const activeCat = document.querySelector(".category-chip.active").dataset.category;
  const list = document.getElementById("formula-list");

  const filtered = FORMULA_DATA.filter((f) => {
    if (activeCat !== "all" && f.category !== activeCat) return false;
    if (!search) return true;
    if (f.name.toLowerCase().includes(search)) return true;
    if (f.category.toLowerCase().includes(search)) return true;
    if (f.desc.toLowerCase().includes(search)) return true;
    for (const v of f.vars) {
      if (v[0].toLowerCase().includes(search) || v[1].toLowerCase().includes(search)) return true;
    }
    return false;
  });

  let html = "";
  filtered.forEach((f) => {
    const sel = selectedFormulaId === f.id ? " selected" : "";
    const varsHtml = f.vars.map((v) => `<span>${v[0]}</span>: ${v[1]}`).join(" &nbsp;|&nbsp; ");
    html += `<div class="formula-card${sel}" data-id="${f.id}">
      <div class="fc-name">${f.name}</div>
      <div class="fc-category">${f.category}</div>
      <div class="fc-display">${renderFormulaPreviewHtml(f.tokens)}</div>
      <div class="fc-description">${f.desc}</div>
      <div class="fc-variables">${varsHtml}</div>
    </div>`;
  });

  if (filtered.length === 0) {
    html = '<div style="text-align:center;color:#999;padding:20px;">No formulas found</div>';
  }

  list.innerHTML = html;

  list.querySelectorAll(".formula-card").forEach((card) => {
    card.addEventListener("click", () => {
      const id = card.dataset.id;
      if (selectedFormulaId === id) {
        selectedFormulaId = null;
        card.classList.remove("selected");
      } else {
        list.querySelectorAll(".formula-card").forEach((c) => c.classList.remove("selected"));
        selectedFormulaId = id;
        card.classList.add("selected");
      }
    });
  });
}

function renderFormulaPreviewHtml(tokens) {
  return tokens.map((tok) => {
    if (tok.y === "f") {
      const num = renderFormulaPreviewHtml(tok.num);
      const den = renderFormulaPreviewHtml(tok.den);
      return `<span style="display:inline-flex;flex-direction:column;align-items:center;vertical-align:middle;margin:0 2px;line-height:1.2"><span style="border-bottom:1px solid #333;padding:0 3px">${num}</span><span style="padding:0 3px">${den}</span></span>`;
    }
    const escaped = tok.t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    switch (tok.y) {
      case "b": return `<sub>${escaped}</sub>`;
      case "p": return `<sup>${escaped}</sup>`;
      case "i": return `<em>${escaped}</em>`;
      default: return escaped;
    }
  }).join("");
}

function renderOmmlTokens(tokens) {
  return tokens.map((tok) => {
    if (tok.y === "f") {
      return `<m:f><m:num>${renderOmmlTokens(tok.num)}</m:num><m:den>${renderOmmlTokens(tok.den)}</m:den></m:f>`;
    }
    let rPr = "";
    if (tok.y === "i") rPr = "<m:rPr><m:sty m:val=\"i\"/></m:rPr>";
    if (tok.y === "b") rPr = "<m:rPr><m:sty m:val=\"p\"/></m:rPr>";
    if (tok.y === "p") rPr = "<m:rPr><m:sty m:val=\"p\"/></m:rPr>";

    let wRPr = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
    if (tok.y === "b") wRPr = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:vertAlign w:val="subscript"/></w:rPr>';
    if (tok.y === "p") wRPr = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:vertAlign w:val="superscript"/></w:rPr>';

    const escaped = tok.t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    return `<m:r>${rPr}${wRPr}<m:t>${escaped}</m:t></m:r>`;
  }).join("");
}

function buildFormulaOoxml(formula) {
  const ommlContent = renderOmmlTokens(formula.tokens);
  const mathPara = `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>${ommlContent}</m:oMath></m:oMathPara></w:p>`;

  let paragraphs = "";
  paragraphs += `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
  paragraphs += mathPara;
  paragraphs += `<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;

  return wrapInOoxml(paragraphs);
}

async function insertSelectedFormula() {
  if (!selectedFormulaId) {
    showStatus("Please select a formula first.", "error");
    return;
  }

  const formula = FORMULA_DATA.find((f) => f.id === selectedFormulaId);
  if (!formula) return;

  try {
    await Word.run(async (context) => {
      const ooxml = buildFormulaOoxml(formula);
      const range = context.document.getSelection();
      const inserted = range.insertOoxml(ooxml, Word.InsertLocation.replace);
      const afterRange = inserted.getRange("After");
      afterRange.select();
      await context.sync();
      showStatus("Formula inserted!", "success");
      document.activeElement.blur();
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
  }
}
