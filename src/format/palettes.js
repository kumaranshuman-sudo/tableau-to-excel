/* Tableau built-in and automatic colour palettes. */

"use strict";

// ── TABLEAU BUILT-IN PALETTE LOOKUP ──────────────────────────────────────
export const TABLEAU_BUILTIN_PALETTES = {
  "blue_10_0": ["#C7DDEA","#AFCFE1","#97C0D7","#7FAFCA","#689BC0","#5487B1","#4475A0","#356790","#2A5783"],
  "orange_10_0": ["#F3C184", "#F0AE62", "#EE9A42", "#EF882D", "#ED7420", "#E25F1D", "#CC531F", "#B54820", "#9E3D22"],
  "green_10_0": ["#B9D9AF", "#A3CF95", "#8BC57D", "#74BA67", "#5DAA56", "#4D984B", "#3D8743", "#31773F", "#24693D"],
  "red_10_0": ["#F3B8AB", "#EEA08E", "#EA8972", "#E9725B", "#EA5C4C", "#E6453C", "#D92C34", "#C71532", "#AE123A"],
  "purple_10_0": ["#ECC6E3", "#E2BCD8", "#D7AFCA", "#CB9FBC", "#BB85A8", "#AC759B", "#9F6B93", "#8E5B86", "#7C4D79"],
  "brown_10_0": ["#E8D5B4","#DEBE8A","#D8A66B","#D08F57","#C97A4B","#BF6740","#B45539","#AA4435","#9F3632"],
  "gray_10_0": ["#E5E5E5","#D4D6D8","#C0C4C8","#AAB0B6","#959DA5","#818A94","#6D7782","#5B6470","#49525E"],
  "gray_warm_10_0": ["#D8D1CE","#CBC2BE","#BBB0AB","#AB9F9A","#9A8E89","#887C77","#776B67","#685D59","#59504E"],
  "blue_teal_10_0": ["#B7D7D1","#9BC9C7","#80BBC0","#67ACC0","#529DBA","#448DAF","#3A7D9F","#336C91","#2C5985"],
  "orange_gold_10_0": ["#E8C85E","#EDB657","#F0A54A","#F08F32","#EF791F","#E96418","#D5531D","#BA4522","#9E3A26"],
  "green_gold_10_0": ["#E5C75A","#CDBE58","#B1B953","#97B64F","#7FB255","#67A957","#529B53","#348347","#146C36"],
  "red_gold_10_0": ["#E8C85A","#EDA951","#F08C4B","#EC7247","#E65E47","#DF4D47","#D33A45","#C32942","#B71D3E"],
  "orange_blue_diverging_10_0": ["#9E3D22","#B54820","#CC531F","#E25F1D","#ED7420","#F3D9BE","#97C0D7","#689BC0","#4475A0","#356790","#2B5C8A"],
  "red_green_diverging_10_0": ["#AE123A","#C61E3F","#D93443","#E94F4A","#F07A66","#F2B3A3","#8BC97D","#6DB65F","#539F50","#3B8447","#24693D"],
  "green_blue_diverging_10_0": ["#24693D","#347D46","#4E9854","#72B464","#9BCF89","#D4DDD9","#A9C9DC","#7DAACE","#5C8FBC","#4373A0","#2A5783"],
  "red_blue_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#E7E2DE","#B8D1E1","#86B0D1","#5F90BC","#44709C","#2E5A87"],
  "red_black_10_0": ["#AE123A","#C71F3E","#DC393F","#EB5A4F","#F3A091","#DDD9D5","#B9BCBC","#969DA1","#78818A","#606A75","#49525E"],
  "gold_purple_diverging_10_0": ["#AD9024","#B89B34","#C6AA50","#D3BA6D","#DDC892","#E3D7D1","#D7C1D2","#C9A5C3","#BB8AB2","#AC7299"],
  "red_green_gold_diverging_10_0": ["#BE2A3E","#D44344","#E75D49","#F07A47","#F2A14A","#E7C65A","#A8BE5E","#77AF5B","#55994E","#3B8746","#22763F"],
  "sunrise_sunset_diverging_10_0": ["#33608C","#556AA0","#7B67A6","#A664A2","#CD6C95","#EC7C79","#F2B15A","#ED8D46","#E56B44","#D24844","#B81840"],
  "orange_blue_white_diverging_10_0": ["#9E3D22","#B94B20","#D45A1D","#EC7420","#F3B562","#F5F1EC","#C6DDEA","#93BED8","#679BC1","#4677A5","#2B5C8A"],
  "red_green_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F1EE","#B9DFAF","#8BC57D","#5DAA56","#3F8847","#24693D"],
  "green_blue_white_diverging_10_0": ["#24693D","#3E864B","#5BA557","#82C06F","#B7DFAE","#F2F3F1","#C7DDEA","#97C0D7","#689BC0","#4475A0","#2A5783"],
  "red_blue_white_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#F4F3F2","#C6DDEA","#93BED8","#679BC1","#4677A5","#2E5A87"],
  "red_black_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F2F1","#D0D2D1","#A7ADB0","#838C93","#626C77","#49525E"],
  "tableau-blue-light": ["#EEF2F7","#E7EDF5","#DFE8F3","#D8E2F0","#D1DDEE","#CBD9ED","#C9D7F1","#C7D9F2","#C4D8F3"],
  "tableau-orange-light": ["#F6F2EE","#F7EBDD","#F8E3CC","#F9DBBE","#FAD5B3","#FBCFA8","#FCCB9F","#FFCC9E","#FFCC9E"],
  "tableau-orange-blue-light": ["#FFCC9E","#FAD1AB","#F4D8BC","#EEE0CC","#EAE4D8","#E8E8E8","#DFE7EF","#D5E0EC","#CBD9E9","#C7D9F1","#C4D8F3"],
  "tableau-map-blue-green": ["#F5F5C8","#EEF2B3","#E0EA9A","#CBE18F","#AFD695","#91CC9D","#72C3A8","#58BCB5","#41B7C4"],
  "tableau-map-temperatur": ["#529985","#669C76","#81A364","#A6B04E","#D2C63F","#F0D347","#F2C04A","#E7A24A","#D4824D","#C26B51"]
};

// ── HELPER: Get built-in palette colors by name ──────────────────────────
export function getBuiltInPaletteColors(paletteName) {
  if (!paletteName) return null;
  
  if (TABLEAU_BUILTIN_PALETTES[paletteName]) {
    return TABLEAU_BUILTIN_PALETTES[paletteName];
  }
  
  const lowerName = paletteName.toLowerCase();
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    if (key.toLowerCase() === lowerName) {
      return colors;
    }
  }
  
  const normalized = paletteName.toLowerCase().replace(/[-\s]+/g, '');
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    const keyNormalized = key.toLowerCase().replace(/[-\s]+/g, '');
    if (keyNormalized.includes(normalized) || normalized.includes(keyNormalized)) {
      return colors;
    }
  }
  
  return null;
}

export const AUTOMATIC_PALETTE_BY_MARK = {
  automatic:    "blue_teal_10_0",
  bar:          "blue_10_0",
  line:         "blue_10_0",
  area:         "blue_teal_10_0",
  square:       "blue_teal_10_0",
  circle:       "blue_10_0",
  shape:        "blue_10_0",
  text:         "blue_10_0",
  map:          "blue_teal_10_0",
  multipolygon: "blue_teal_10_0",
  pie:          "blue_10_0",
  ganttbar:     "blue_10_0",
  polygon:      "blue_10_0",
  density:      "blue_10_0",
  heatmap:      "blue_10_0"
};

export const DEFAULT_AUTOMATIC_PALETTE = "blue_10_0";

export function getAutomaticPaletteForMark(markClass) {
  const key = String(markClass || "Automatic").toLowerCase().replace(/[\s_-]+/g, "");
  const name = AUTOMATIC_PALETTE_BY_MARK[key] || DEFAULT_AUTOMATIC_PALETTE;
  if (!AUTOMATIC_PALETTE_BY_MARK[key]) {
    console.log(`[Auto Palette] Mark type "${markClass}" not in table, using ${name}`);
  }
  return {
    name,
    colors: TABLEAU_BUILTIN_PALETTES[name].map(c => "FF" + c.replace("#", "").toUpperCase())
  };
}

export const TABLEAU_10 = ["FF4E79A7", "FFF28E2B", "FFE15759", "FF76B7B2", "FF59A14F",
                    "FFEDC948", "FFB07AA1", "FFFF9DA7", "FF9C755F", "FFBAB0AC"];
