import { useEffect, useRef, useState, useCallback, useLayoutEffect } from 'react'
import { supabase } from '../lib/supabase'

// ── Attribution: Base anatomical references by Servier Medical Art (smart.servier.com) CC BY 4.0 ──

// ── Inline professional SVG body templates ─────────────────────────────────────────────────────────
// Self-contained — no Supabase Storage dependency, works offline, loads instantly

const SKIN = '#F4C5A3'
const SKIN2 = '#E8A87C'
const SKIN_STROKE = '#8B5A3C'
const HAIR = '#3D2314'
const FEATURE = '#C4785A'
const LANDMARK = '#B86A4A'
const GUIDE = '#D4A088'
const REF_LINE = 'rgba(180,100,70,0.3)'

const svgFemaleBodyFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 500">
<defs>
  <linearGradient id="sk" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- hair -->
<ellipse cx="130" cy="32" rx="34" ry="44" fill="${HAIR}"/>
<path d="M96,18 Q130,-8 164,18" stroke="${HAIR}" stroke-width="3" fill="none" stroke-linecap="round"/>
<!-- head -->
<ellipse cx="130" cy="40" rx="30" ry="36" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- eyes -->
<ellipse cx="117" cy="35" rx="5" ry="3.5" fill="none" stroke="${FEATURE}" stroke-width="1"/>
<circle cx="117" cy="35" r="2" fill="#2C1A0E"/>
<ellipse cx="143" cy="35" rx="5" ry="3.5" fill="none" stroke="${FEATURE}" stroke-width="1"/>
<circle cx="143" cy="35" r="2" fill="#2C1A0E"/>
<!-- eyebrows -->
<path d="M111,29 Q117,26 123,29" stroke="${HAIR}" stroke-width="1.2" fill="none" stroke-linecap="round"/>
<path d="M137,29 Q143,26 149,29" stroke="${HAIR}" stroke-width="1.2" fill="none" stroke-linecap="round"/>
<!-- nose -->
<path d="M130,38 L127,46 Q130,48 133,46 L130,38" stroke="${FEATURE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<!-- lips -->
<path d="M123,52 Q130,50 137,52 Q130,57 123,52Z" fill="#D4826A" stroke="${FEATURE}" stroke-width="0.8"/>
<!-- neck -->
<path d="M116,74 L110,100 L150,100 L144,74Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left shoulder/arm -->
<path d="M62,124 C54,132 46,152 44,174 C42,192 44,208 48,222 L58,232 L68,226 C65,210 63,192 64,172 C65,152 70,134 72,126Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left forearm -->
<path d="M48,222 C42,236 38,252 38,268 C38,280 42,288 48,292 L58,284 C54,280 52,272 52,262 C52,250 56,236 60,226Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left hand -->
<path d="M38,270 C34,276 34,284 38,290 C42,296 48,298 54,296 C58,293 60,288 58,282 C56,277 52,272 48,270Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M32,268 C30,262 31,256 34,254 C37,252 39,256 38,260" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<path d="M38,265 C36,259 37,253 40,251 C43,249 45,253 44,257" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<path d="M44,267 C43,261 44,255 47,253 C50,251 52,255 51,259" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<!-- right shoulder/arm -->
<path d="M198,124 C206,132 214,152 216,174 C218,192 216,208 212,222 L202,232 L192,226 C195,210 197,192 196,172 C195,152 190,134 188,126Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right forearm -->
<path d="M212,222 C218,236 222,252 222,268 C222,280 218,288 212,292 L202,284 C206,280 208,272 208,262 C208,250 204,236 200,226Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right hand -->
<path d="M222,270 C226,276 226,284 222,290 C218,296 212,298 206,296 C202,293 200,288 202,282 C204,277 208,272 212,270Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M228,268 C230,262 229,256 226,254 C223,252 221,256 222,260" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<path d="M222,265 C224,259 223,253 220,251 C217,249 215,253 216,257" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<path d="M216,267 C217,261 216,255 213,253 C210,251 208,255 209,259" stroke="${SKIN_STROKE}" stroke-width="0.9" fill="none" stroke-linecap="round"/>
<!-- torso -->
<path d="M72,126 C78,116 100,106 116,100 L144,100 C160,106 182,116 188,126 C204,138 212,158 212,182 C212,206 206,226 202,242 C198,256 192,268 188,278 L180,290 C174,298 168,304 160,306 L100,306 C92,304 86,298 80,290 L72,278 C68,268 62,256 58,242 C54,226 48,206 48,182 C48,158 56,138 72,126Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- breast contour lines -->
<path d="M78,152 Q104,188 130,183 Q156,188 182,152" stroke="${GUIDE}" stroke-width="1.3" fill="none" opacity="0.8"/>
<path d="M76,148 C88,162 105,174 114,178" stroke="${GUIDE}" stroke-width="0.9" fill="none" opacity="0.5"/>
<path d="M184,148 C172,162 155,174 146,178" stroke="${GUIDE}" stroke-width="0.9" fill="none" opacity="0.5"/>
<!-- inframammary fold -->
<path d="M80,192 Q104,204 130,200 Q156,204 180,192" stroke="${LANDMARK}" stroke-width="1" fill="none" stroke-dasharray="3,2" opacity="0.6"/>
<!-- nipple markers -->
<circle cx="108" cy="174" r="4" fill="${FEATURE}" opacity="0.85"/>
<circle cx="152" cy="174" r="4" fill="${FEATURE}" opacity="0.85"/>
<!-- midline reference -->
<line x1="130" y1="100" x2="130" y2="306" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="5,4"/>
<!-- umbilicus -->
<ellipse cx="130" cy="234" rx="5.5" ry="6" fill="none" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<ellipse cx="130" cy="236" rx="3" ry="3.5" fill="${FEATURE}" opacity="0.45"/>
<!-- suprapubic -->
<path d="M106,296 Q130,290 154,296" stroke="${GUIDE}" stroke-width="1" fill="none" opacity="0.5"/>
<!-- left leg -->
<path d="M100,306 C90,312 80,330 76,356 C72,378 72,398 74,416 C76,430 80,446 82,460 L94,468 L106,464 C108,450 112,434 114,418 C116,400 116,380 114,357 C112,332 106,314 104,306Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left knee -->
<ellipse cx="88" cy="438" rx="14" ry="18" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<!-- left lower leg -->
<path d="M82,460 C78,474 76,486 76,494 C76,498 79,500 84,500 L100,500 L104,496 C104,488 102,474 100,462 L94,468Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right leg -->
<path d="M160,306 C170,312 180,330 184,356 C188,378 188,398 186,416 C184,430 180,446 178,460 L166,468 L154,464 C152,450 148,434 146,418 C144,400 144,380 146,357 C148,332 154,314 156,306Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right knee -->
<ellipse cx="172" cy="438" rx="14" ry="18" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<!-- right lower leg -->
<path d="M178,460 C182,474 184,486 184,494 C184,498 181,500 176,500 L160,500 L156,496 C156,488 158,474 160,462 L166,468Z" fill="url(#sk)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
</svg>`

const svgFemaleBodyBack = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 500">
<defs>
  <linearGradient id="sk2" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- hair (longer in back) -->
<ellipse cx="130" cy="28" rx="34" ry="48" fill="${HAIR}"/>
<path d="M96,70 Q80,120 84,160" stroke="${HAIR}" stroke-width="4" fill="none" stroke-linecap="round"/>
<path d="M164,70 Q180,120 176,160" stroke="${HAIR}" stroke-width="4" fill="none" stroke-linecap="round"/>
<!-- head -->
<ellipse cx="130" cy="40" rx="30" ry="36" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- neck -->
<path d="M114,74 L108,100 L152,100 L146,74Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left arm -->
<path d="M60,124 C52,132 44,152 42,174 C40,192 42,208 46,222 L56,232 L66,226 C63,210 61,192 62,172 C63,152 68,134 70,126Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M46,222 C40,236 36,252 36,268 C36,280 40,288 46,292 L56,284 C52,280 50,272 50,262 C50,250 54,236 58,226Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M36,272 C32,278 32,286 36,292 C40,298 46,300 52,298 C56,295 58,290 56,284Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right arm -->
<path d="M200,124 C208,132 216,152 218,174 C220,192 218,208 214,222 L204,232 L194,226 C197,210 199,192 198,172 C197,152 192,134 190,126Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M214,222 C220,236 224,252 224,268 C224,280 220,288 214,292 L204,284 C208,280 210,272 210,262 C210,250 206,236 202,226Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M224,272 C228,278 228,286 224,292 C220,298 214,300 208,298 C204,295 202,290 204,284Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- torso back -->
<path d="M70,126 C76,116 100,106 114,100 L146,100 C160,106 184,116 190,126 C206,138 214,158 214,182 C214,206 208,226 204,242 C200,256 194,268 190,278 L182,290 C176,298 170,304 162,306 L98,306 C90,304 84,298 78,290 L70,278 C66,268 60,256 56,242 C52,226 46,206 46,182 C46,158 54,138 70,126Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- spine reference line -->
<line x1="130" y1="100" x2="130" y2="306" stroke="${REF_LINE}" stroke-width="0.9" stroke-dasharray="5,4"/>
<!-- shoulder blade indicators -->
<ellipse cx="102" cy="150" rx="20" ry="28" fill="none" stroke="${GUIDE}" stroke-width="1" opacity="0.5" transform="rotate(-10 102 150)"/>
<ellipse cx="158" cy="150" rx="20" ry="28" fill="none" stroke="${GUIDE}" stroke-width="1" opacity="0.5" transform="rotate(10 158 150)"/>
<!-- lumbar dimples -->
<circle cx="116" cy="268" r="5" fill="none" stroke="${LANDMARK}" stroke-width="1.2" opacity="0.7"/>
<circle cx="144" cy="268" r="5" fill="none" stroke="${LANDMARK}" stroke-width="1.2" opacity="0.7"/>
<!-- gluteal fold -->
<path d="M86,300 Q130,310 174,300" stroke="${GUIDE}" stroke-width="1.2" fill="none" opacity="0.6"/>
<!-- left leg back -->
<path d="M98,306 C88,312 78,330 74,356 C70,378 70,398 72,416 C74,430 78,446 80,460 L92,468 L104,464 C106,450 110,434 112,418 C114,400 114,380 112,357 C110,332 104,314 102,306Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<ellipse cx="86" cy="438" rx="14" ry="18" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<path d="M80,460 C76,474 74,486 74,494 C74,498 77,500 82,500 L98,500 L102,496 C102,488 100,474 98,462 L92,468Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right leg back -->
<path d="M162,306 C172,312 182,330 186,356 C190,378 190,398 188,416 C186,430 182,446 180,460 L168,468 L156,464 C154,450 150,434 148,418 C146,400 146,380 148,357 C150,332 156,314 158,306Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<ellipse cx="174" cy="438" rx="14" ry="18" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<path d="M180,460 C184,474 186,486 186,494 C186,498 183,500 178,500 L162,500 L158,496 C158,488 160,474 162,462 L168,468Z" fill="url(#sk2)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
</svg>`

const svgBreastFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 280">
<defs>
  <linearGradient id="sk3" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
  <radialGradient id="bgrad" cx="50%" cy="35%" r="60%">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </radialGradient>
</defs>
<!-- neck -->
<path d="M118,8 L112,40 L188,40 L182,8Z" fill="url(#sk3)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- clavicle area -->
<path d="M112,40 Q70,48 40,68" stroke="${GUIDE}" stroke-width="1.5" fill="none"/>
<path d="M188,40 Q230,48 260,68" stroke="${GUIDE}" stroke-width="1.5" fill="none"/>
<!-- left clavicle bone indicator -->
<path d="M112,40 Q80,46 52,58" stroke="${LANDMARK}" stroke-width="2" fill="none" opacity="0.6"/>
<path d="M188,40 Q220,46 248,58" stroke="${LANDMARK}" stroke-width="2" fill="none" opacity="0.6"/>
<!-- chest wall / ribcage outline -->
<path d="M40,68 C38,90 36,115 38,140 C40,165 46,185 52,200 C58,215 66,225 74,232 L100,248 L150,252 L200,248 L226,232 C234,225 242,215 248,200 C254,185 260,165 262,140 C264,115 262,90 260,68Z" fill="url(#sk3)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- midline reference -->
<line x1="150" y1="40" x2="150" y2="270" stroke="${REF_LINE}" stroke-width="1" stroke-dasharray="5,4"/>
<!-- breast meridian lines (surgical planning) -->
<line x1="96" y1="48" x2="96" y2="248" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="3,4" opacity="0.7"/>
<line x1="204" y1="48" x2="204" y2="248" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="3,4" opacity="0.7"/>
<!-- sternal notch marker -->
<circle cx="150" cy="44" r="5" fill="none" stroke="${LANDMARK}" stroke-width="1.5" opacity="0.8"/>
<!-- left breast contour -->
<path d="M52,80 C48,100 46,125 50,148 C54,168 62,184 72,196 C82,206 92,212 100,215 C108,218 115,218 120,215 C126,212 132,205 136,196 C140,185 142,170 140,155 C138,138 132,120 124,107 C116,94 106,86 96,82 C86,78 72,76 62,78Z" fill="url(#bgrad)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- right breast contour -->
<path d="M248,80 C252,100 254,125 250,148 C246,168 238,184 228,196 C218,206 208,212 200,215 C192,218 185,218 180,215 C174,212 168,205 164,196 C160,185 158,170 160,155 C162,138 168,120 176,107 C184,94 194,86 204,82 C214,78 228,76 238,78Z" fill="url(#bgrad)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- inframammary fold — critical surgical landmark -->
<path d="M50,210 Q90,228 120,224 Q136,222 140,218" stroke="${LANDMARK}" stroke-width="2" fill="none" stroke-dasharray="4,2" opacity="0.9"/>
<path d="M250,210 Q210,228 180,224 Q164,222 160,218" stroke="${LANDMARK}" stroke-width="2" fill="none" stroke-dasharray="4,2" opacity="0.9"/>
<!-- left NAC (nipple-areola complex) -->
<circle cx="96" cy="165" r="18" fill="none" stroke="${FEATURE}" stroke-width="1.2" opacity="0.6"/>
<circle cx="96" cy="165" r="9" fill="none" stroke="${FEATURE}" stroke-width="1.5" opacity="0.8"/>
<circle cx="96" cy="165" r="4" fill="${FEATURE}" opacity="0.9"/>
<!-- right NAC -->
<circle cx="204" cy="165" r="18" fill="none" stroke="${FEATURE}" stroke-width="1.2" opacity="0.6"/>
<circle cx="204" cy="165" r="9" fill="none" stroke="${FEATURE}" stroke-width="1.5" opacity="0.8"/>
<circle cx="204" cy="165" r="4" fill="${FEATURE}" opacity="0.9"/>
<!-- anterior axillary line markers -->
<line x1="44" y1="68" x2="44" y2="240" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4" opacity="0.5"/>
<line x1="256" y1="68" x2="256" y2="240" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4" opacity="0.5"/>
<!-- IMF label zone -->
<text x="12" y="220" font-size="8" fill="${LANDMARK}" opacity="0.7" font-family="Arial,sans-serif">IMF</text>
<text x="272" y="220" font-size="8" fill="${LANDMARK}" opacity="0.7" font-family="Arial,sans-serif">IMF</text>
<!-- midline label -->
<text x="153" y="260" font-size="8" fill="${GUIDE}" opacity="0.6" font-family="Arial,sans-serif">ML</text>
</svg>`

const svgFaceFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 320">
<defs>
  <linearGradient id="sk4" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- hair -->
<ellipse cx="120" cy="80" rx="86" ry="100" fill="${HAIR}"/>
<path d="M34,80 Q30,140 36,180" stroke="${HAIR}" stroke-width="6" fill="none" stroke-linecap="round"/>
<path d="M206,80 Q210,140 204,180" stroke="${HAIR}" stroke-width="6" fill="none" stroke-linecap="round"/>
<!-- face -->
<ellipse cx="120" cy="148" rx="78" ry="92" fill="url(#sk4)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- facial thirds reference lines (surgical planning) -->
<line x1="42" y1="76" x2="198" y2="76" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
<line x1="42" y1="118" x2="198" y2="118" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
<line x1="42" y1="162" x2="198" y2="162" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
<line x1="42" y1="220" x2="198" y2="220" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
<!-- midline -->
<line x1="120" y1="56" x2="120" y2="240" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="5,4"/>
<!-- forehead -->
<path d="M60,100 Q120,90 180,100" stroke="${GUIDE}" stroke-width="0.8" fill="none" opacity="0.4"/>
<!-- left eyebrow -->
<path d="M68,116 Q85,109 102,114" stroke="${HAIR}" stroke-width="2.5" fill="none" stroke-linecap="round"/>
<!-- right eyebrow -->
<path d="M138,114 Q155,109 172,116" stroke="${HAIR}" stroke-width="2.5" fill="none" stroke-linecap="round"/>
<!-- left eye -->
<path d="M68,126 Q85,118 102,126 Q85,134 68,126Z" fill="#FDEBD8" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<ellipse cx="85" cy="126" rx="7" ry="6" fill="#5C3317"/>
<circle cx="85" cy="124" r="4" fill="#1A0E08"/>
<circle cx="88" cy="122" r="1.5" fill="white"/>
<!-- right eye -->
<path d="M138,126 Q155,118 172,126 Q155,134 138,126Z" fill="#FDEBD8" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<ellipse cx="155" cy="126" rx="7" ry="6" fill="#5C3317"/>
<circle cx="155" cy="124" r="4" fill="#1A0E08"/>
<circle cx="158" cy="122" r="1.5" fill="white"/>
<!-- nose bridge -->
<path d="M116,126 L112,152 L108,162" stroke="${FEATURE}" stroke-width="1" fill="none" opacity="0.6"/>
<path d="M124,126 L128,152 L132,162" stroke="${FEATURE}" stroke-width="1" fill="none" opacity="0.6"/>
<!-- nose tip -->
<path d="M108,162 Q120,170 132,162 Q124,175 120,174 Q116,175 108,162Z" fill="${SKIN2}" stroke="${SKIN_STROKE}" stroke-width="1"/>
<!-- nostrils -->
<ellipse cx="112" cy="167" rx="6" ry="4" fill="${FEATURE}" opacity="0.5"/>
<ellipse cx="128" cy="167" rx="6" ry="4" fill="${FEATURE}" opacity="0.5"/>
<!-- philtrum -->
<path d="M116,172 Q120,176 124,172 L122,185 Q120,186 118,185Z" fill="${SKIN2}" stroke="${SKIN_STROKE}" stroke-width="0.8"/>
<!-- lips -->
<path d="M96,190 Q108,182 120,186 Q132,182 144,190 Q132,186 120,188 Q108,186 96,190Z" fill="#D4826A" stroke="${FEATURE}" stroke-width="1"/>
<path d="M96,190 Q100,200 108,204 Q120,207 132,204 Q140,200 144,190 Q132,196 120,197 Q108,196 96,190Z" fill="#C4785A" stroke="${FEATURE}" stroke-width="1"/>
<!-- chin -->
<path d="M94,210 Q120,234 146,210 Q132,226 120,228 Q108,226 94,210Z" fill="url(#sk4)" stroke="${SKIN_STROKE}" stroke-width="0.8"/>
<!-- ear left -->
<path d="M42,128 C34,134 30,148 32,162 C34,176 40,184 46,184 C48,180 48,172 46,162 C44,150 44,138 46,130Z" fill="url(#sk4)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- ear right -->
<path d="M198,128 C206,134 210,148 208,162 C206,176 200,184 194,184 C192,180 192,172 194,162 C196,150 196,138 194,130Z" fill="url(#sk4)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- neck -->
<path d="M72,234 L66,295 L174,295 L168,234 C154,244 86,244 72,234Z" fill="url(#sk4)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- sternocleidomastoid indicators -->
<path d="M66,295 Q80,268 88,245" stroke="${GUIDE}" stroke-width="1" fill="none" opacity="0.4"/>
<path d="M174,295 Q160,268 152,245" stroke="${GUIDE}" stroke-width="1" fill="none" opacity="0.4"/>
<!-- facial zone labels -->
<text x="6" y="80" font-size="7" fill="${GUIDE}" opacity="0.6" font-family="Arial,sans-serif">III</text>
<text x="6" y="122" font-size="7" fill="${GUIDE}" opacity="0.6" font-family="Arial,sans-serif">II</text>
<text x="6" y="166" font-size="7" fill="${GUIDE}" opacity="0.6" font-family="Arial,sans-serif">I</text>
</svg>`

const svgFaceLateral = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 320">
<defs>
  <linearGradient id="sk5" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0%" stop-color="${SKIN2}"/>
    <stop offset="100%" stop-color="${SKIN}"/>
  </linearGradient>
</defs>
<!-- hair (back) -->
<path d="M130,20 C160,15 190,20 200,40 C210,60 208,90 200,120 C192,150 180,180 176,210 C172,240 174,270 170,290 L130,295 L134,250 C138,220 144,190 146,160 C148,130 146,100 142,70 C138,45 132,28 130,20Z" fill="${HAIR}"/>
<!-- face profile -->
<path d="M130,30 C110,32 95,40 85,55 C75,70 72,88 70,108 C68,128 68,150 72,168 C76,184 83,196 90,206 C97,216 106,220 112,220 C108,228 105,240 108,250 C112,260 122,266 132,264 C140,262 148,256 150,248 C152,240 148,230 140,224 C148,216 156,204 162,190 C168,174 172,155 170,136 C168,116 162,98 154,82 C146,66 138,48 130,30Z" fill="url(#sk5)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- forehead detail -->
<path d="M130,30 C126,44 124,62 126,80" stroke="${GUIDE}" stroke-width="0.8" fill="none" opacity="0.4"/>
<!-- nasal dorsum -->
<path d="M100,96 L92,114 L90,130 L86,148 L88,158" stroke="${SKIN_STROKE}" stroke-width="1.5" fill="none" stroke-linecap="round"/>
<!-- nasal tip -->
<path d="M88,158 Q82,165 85,170 Q90,176 97,175 Q102,170 100,163" stroke="${SKIN_STROKE}" stroke-width="1.2" fill="none"/>
<path d="M85,170 L80,175" stroke="${SKIN_STROKE}" stroke-width="1" fill="none" stroke-linecap="round"/>
<!-- eye -->
<path d="M108,110 Q118,104 128,110 Q118,118 108,110Z" fill="#FDEBD8" stroke="${SKIN_STROKE}" stroke-width="1"/>
<circle cx="120" cy="110" r="5" fill="#2C1A0E"/>
<circle cx="122" cy="108" r="1.2" fill="white"/>
<!-- eyebrow -->
<path d="M106,102 Q118,96 130,100" stroke="${HAIR}" stroke-width="2" fill="none" stroke-linecap="round"/>
<!-- eyelid crease -->
<path d="M108,106 Q118,102 128,106" stroke="${SKIN_STROKE}" stroke-width="0.8" fill="none" opacity="0.5"/>
<!-- ear -->
<path d="M152,120 C160,125 165,138 163,152 C161,164 155,172 148,173 C144,170 142,162 143,150 C144,138 148,126 152,120Z" fill="url(#sk5)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<path d="M155,132 C158,138 158,148 155,155" stroke="${SKIN_STROKE}" stroke-width="1" fill="none"/>
<path d="M151,138 C152,143 151,150 149,154" stroke="${SKIN_STROKE}" stroke-width="0.8" fill="none"/>
<!-- lips lateral -->
<path d="M90,194 Q95,188 102,190 Q108,192 110,196 Q106,202 100,204 Q94,202 90,194Z" fill="#D4826A" stroke="${FEATURE}" stroke-width="1"/>
<!-- chin -->
<path d="M108,220 Q112,235 110,248 Q112,258 122,262 Q130,262 135,256" stroke="${SKIN_STROKE}" stroke-width="1.2" fill="none"/>
<!-- nasolabial fold -->
<path d="M88,166 Q90,178 91,190" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="2,2" opacity="0.6"/>
<!-- neck lateral -->
<path d="M112,220 C108,230 106,250 108,268 C110,280 114,290 120,295" stroke="${SKIN_STROKE}" stroke-width="1.2" fill="none"/>
<path d="M140,224 C144,236 148,252 146,268 C144,280 140,290 132,295" stroke="${SKIN_STROKE}" stroke-width="1.2" fill="none"/>
<!-- Frankfort horizontal reference -->
<line x1="60" y1="108" x2="165" y2="108" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
<!-- vertical reference -->
<line x1="78" y1="30" x2="78" y2="290" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="4,4"/>
</svg>`

const svgAbdomenFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 280 300">
<defs>
  <linearGradient id="sk6" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- ribcage outline (top) -->
<path d="M40,20 C38,40 36,65 40,85 C44,100 54,110 68,114 L100,120 L140,120 L180,120 L212,114 C226,110 236,100 240,85 C244,65 242,40 240,20Z" fill="url(#sk6)" stroke="${SKIN_STROKE}" stroke-width="1.3" opacity="0.7"/>
<!-- ribcage costal margin -->
<path d="M68,114 Q100,128 140,126 Q180,128 212,114" stroke="${SKIN_STROKE}" stroke-width="1.5" fill="none"/>
<!-- main abdominal wall -->
<path d="M68,114 C56,120 46,135 42,155 C38,175 38,198 42,218 C46,236 54,252 64,262 C74,272 88,278 104,282 L140,285 L176,282 C192,278 206,272 216,262 C226,252 234,236 238,218 C242,198 242,175 238,155 C234,135 224,120 212,114Z" fill="url(#sk6)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- midline (linea alba) — key surgical landmark -->
<line x1="140" y1="115" x2="140" y2="285" stroke="${LANDMARK}" stroke-width="1.5" stroke-dasharray="5,3" opacity="0.8"/>
<!-- rectus abdominis outlines -->
<path d="M115,118 L112,285" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="3,3" opacity="0.5"/>
<path d="M165,118 L168,285" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="3,3" opacity="0.5"/>
<!-- tendinous intersections (3 pairs) -->
<path d="M112,145 Q140,143 168,145" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<path d="M112,175 Q140,173 168,175" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<path d="M112,205 Q140,203 168,205" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<!-- umbilicus — critical surgical landmark -->
<ellipse cx="140" cy="168" rx="10" ry="12" fill="none" stroke="${SKIN_STROKE}" stroke-width="2"/>
<ellipse cx="140" cy="170" rx="6" ry="7" fill="${FEATURE}" opacity="0.55"/>
<circle cx="140" cy="168" r="3" fill="${SKIN_STROKE}" opacity="0.5"/>
<!-- ASIS (anterior superior iliac spine) landmarks -->
<circle cx="56" cy="256" r="7" fill="none" stroke="${LANDMARK}" stroke-width="2" opacity="0.8"/>
<circle cx="224" cy="256" r="7" fill="none" stroke="${LANDMARK}" stroke-width="2" opacity="0.8"/>
<!-- interASIS reference line -->
<line x1="56" y1="256" x2="224" y2="256" stroke="${REF_LINE}" stroke-width="0.8" stroke-dasharray="5,4"/>
<!-- pubic symphysis -->
<ellipse cx="140" cy="280" rx="14" ry="8" fill="${GUIDE}" opacity="0.3" stroke="${LANDMARK}" stroke-width="1.2"/>
<!-- flanks / external oblique outline -->
<path d="M42,155 C36,165 34,180 36,195 C38,210 44,224 52,234 C60,244 70,252 80,256 C90,260 100,262 110,263" stroke="${GUIDE}" stroke-width="1" fill="none" opacity="0.5"/>
<path d="M238,155 C244,165 246,180 244,195 C242,210 236,224 228,234 C220,244 210,252 200,256 C190,260 180,262 170,263" stroke="${GUIDE}" stroke-width="1" fill="none" opacity="0.5"/>
<!-- typical abdominoplasty incision zone (dashed) -->
<path d="M56,272 Q100,285 140,283 Q180,285 224,272" stroke="#E05A2B" stroke-width="1.5" fill="none" stroke-dasharray="6,3" opacity="0.4"/>
<!-- semilunar line / inguinal ligament -->
<path d="M58,254 Q80,268 110,272 Q126,274 140,273" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="3,3" opacity="0.4"/>
<path d="M222,254 Q200,268 170,272 Q154,274 140,273" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="3,3" opacity="0.4"/>
<!-- labels -->
<text x="143" y="162" font-size="7" fill="${LANDMARK}" opacity="0.7" font-family="Arial,sans-serif">UMB</text>
<text x="30" y="260" font-size="7" fill="${LANDMARK}" opacity="0.7" font-family="Arial,sans-serif">ASIS</text>
<text x="228" y="260" font-size="7" fill="${LANDMARK}" opacity="0.7" font-family="Arial,sans-serif">ASIS</text>
<text x="126" y="295" font-size="7" fill="${GUIDE}" opacity="0.6" font-family="Arial,sans-serif">Midline</text>
</svg>`

const svgMaleBodyFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 500">
<defs>
  <linearGradient id="sk7" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- hair (short) -->
<ellipse cx="130" cy="32" rx="32" ry="36" fill="${HAIR}"/>
<!-- head -->
<ellipse cx="130" cy="40" rx="32" ry="37" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- stronger jaw -->
<path d="M98,68 Q106,82 130,86 Q154,82 162,68" stroke="${SKIN_STROKE}" stroke-width="1.2" fill="none"/>
<!-- eyes -->
<ellipse cx="114" cy="35" rx="6" ry="4" fill="none" stroke="${FEATURE}" stroke-width="1.2"/>
<circle cx="114" cy="35" r="2.5" fill="#2C1A0E"/>
<ellipse cx="146" cy="35" rx="6" ry="4" fill="none" stroke="${FEATURE}" stroke-width="1.2"/>
<circle cx="146" cy="35" r="2.5" fill="#2C1A0E"/>
<!-- brows (stronger) -->
<path d="M107,28 Q114,24 121,28" stroke="${HAIR}" stroke-width="2" fill="none" stroke-linecap="round"/>
<path d="M139,28 Q146,24 153,28" stroke="${HAIR}" stroke-width="2" fill="none" stroke-linecap="round"/>
<!-- nose -->
<path d="M130,38 L126,48 Q130,51 134,48 L130,38" stroke="${FEATURE}" stroke-width="1" fill="none"/>
<!-- lips (thinner) -->
<path d="M120,56 Q130,54 140,56 Q130,61 120,56Z" fill="#C4785A" stroke="${FEATURE}" stroke-width="0.8"/>
<!-- neck (wider) -->
<path d="M112,74 L104,102 L156,102 L148,74Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<!-- left arm (larger) -->
<path d="M52,122 C44,130 36,152 34,176 C32,196 34,214 40,228 L52,240 L64,234 C60,218 58,198 60,176 C62,154 68,134 72,124Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<path d="M40,228 C34,244 30,260 30,276 C30,290 34,298 42,302 L54,294 C50,290 48,280 48,268 C48,256 52,242 56,232Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<path d="M30,278 C26,286 26,294 30,300 C34,308 42,312 50,308 C54,305 56,298 54,292Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right arm -->
<path d="M208,122 C216,130 224,152 226,176 C228,196 226,214 220,228 L208,240 L196,234 C200,218 202,198 200,176 C198,154 192,134 188,124Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<path d="M220,228 C226,244 230,260 230,276 C230,290 226,298 218,302 L206,294 C210,290 212,280 212,268 C212,256 208,242 204,232Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<path d="M230,278 C234,286 234,294 230,300 C226,308 218,312 210,308 C206,305 204,298 206,292Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- torso (broader shoulders, narrower hips) -->
<path d="M72,124 C78,114 104,104 116,102 L144,102 C156,104 182,114 188,124 C206,136 216,158 216,184 C216,208 210,230 206,248 C202,262 196,274 190,282 L180,294 C174,300 166,306 158,308 L102,308 C94,306 86,300 80,294 L70,282 C64,274 58,262 54,248 C50,230 44,208 44,184 C44,158 54,136 72,124Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- pectoral muscle outlines -->
<path d="M72,130 Q100,144 128,148 Q132,148 130,130" stroke="${GUIDE}" stroke-width="1.2" fill="none" opacity="0.6"/>
<path d="M188,130 Q160,144 132,148 Q128,148 130,130" stroke="${GUIDE}" stroke-width="1.2" fill="none" opacity="0.6"/>
<!-- nipples (male) -->
<circle cx="104" cy="160" r="3" fill="${FEATURE}" opacity="0.7"/>
<circle cx="156" cy="160" r="3" fill="${FEATURE}" opacity="0.7"/>
<!-- abdominal definition lines -->
<line x1="130" y1="102" x2="130" y2="308" stroke="${REF_LINE}" stroke-width="0.9" stroke-dasharray="5,4"/>
<path d="M106,178 Q130,176 154,178" stroke="${GUIDE}" stroke-width="0.8" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<path d="M104,204 Q130,202 156,204" stroke="${GUIDE}" stroke-width="0.8" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<path d="M104,228 Q130,226 156,228" stroke="${GUIDE}" stroke-width="0.8" fill="none" stroke-dasharray="2,2" opacity="0.4"/>
<!-- umbilicus -->
<ellipse cx="130" cy="228" rx="5" ry="6" fill="none" stroke="${SKIN_STROKE}" stroke-width="1.5"/>
<!-- left leg (wider, straighter) -->
<path d="M102,308 C92,314 80,334 76,360 C72,382 72,402 74,420 C76,436 80,452 82,466 L96,474 L110,470 C112,454 116,438 118,422 C120,404 120,382 118,360 C116,336 108,316 106,308Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<ellipse cx="86" cy="444" rx="15" ry="19" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<path d="M82,466 C78,480 76,492 76,500 L100,500 L106,498 C106,488 104,476 100,468 L96,474Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right leg -->
<path d="M158,308 C168,314 180,334 184,360 C188,382 188,402 186,420 C184,436 180,452 178,466 L164,474 L150,470 C148,454 144,438 142,422 C140,404 140,382 142,360 C144,336 152,316 154,308Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<ellipse cx="174" cy="444" rx="15" ry="19" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1"/>
<path d="M178,466 C182,480 184,492 184,500 L160,500 L154,498 C154,488 156,476 160,468 L164,474Z" fill="url(#sk7)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
</svg>`

const svgLowerBodyFront = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 280 320">
<defs>
  <linearGradient id="sk8" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" stop-color="${SKIN}"/>
    <stop offset="100%" stop-color="${SKIN2}"/>
  </linearGradient>
</defs>
<!-- lower trunk -->
<path d="M44,10 C40,28 38,50 40,70 C42,88 50,104 60,116 C70,128 82,136 94,142 L120,148 L160,148 L186,142 C198,136 210,128 220,116 C230,104 238,88 240,70 C242,50 240,28 236,10Z" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.3"/>
<!-- midline -->
<line x1="140" y1="8" x2="140" y2="148" stroke="${REF_LINE}" stroke-width="0.9" stroke-dasharray="5,4"/>
<!-- inguinal ligament lines -->
<path d="M60,116 Q100,138 140,136 Q180,138 220,116" stroke="${GUIDE}" stroke-width="1" fill="none" stroke-dasharray="4,3" opacity="0.6"/>
<!-- ASIS markers -->
<circle cx="48" cy="74" r="7" fill="none" stroke="${LANDMARK}" stroke-width="2" opacity="0.8"/>
<circle cx="232" cy="74" r="7" fill="none" stroke="${LANDMARK}" stroke-width="2" opacity="0.8"/>
<!-- pubic region -->
<path d="M112,140 Q140,148 168,140" stroke="${SKIN_STROKE}" stroke-width="1" fill="none"/>
<!-- left thigh -->
<path d="M94,142 C80,148 68,162 60,182 C52,200 50,220 54,240 C58,256 68,268 80,276 L104,284 L122,280 C124,264 126,248 126,232 C126,214 122,196 116,180 C110,165 100,152 96,144Z" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right thigh -->
<path d="M186,142 C200,148 212,162 220,182 C228,200 230,220 226,240 C222,256 212,268 200,276 L176,284 L158,280 C156,264 154,248 154,232 C154,214 158,196 164,180 C170,165 180,152 184,144Z" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- inner thigh reference -->
<path d="M122,148 C126,168 130,195 130,222 C130,248 128,272 126,290" stroke="${GUIDE}" stroke-width="0.8" fill="none" stroke-dasharray="3,3" opacity="0.4"/>
<path d="M158,148 C154,168 150,195 150,222 C150,248 152,272 154,290" stroke="${GUIDE}" stroke-width="0.8" fill="none" stroke-dasharray="3,3" opacity="0.4"/>
<!-- left knee -->
<ellipse cx="90" cy="276" rx="22" ry="18" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right knee -->
<ellipse cx="190" cy="276" rx="22" ry="18" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- left lower leg -->
<path d="M80,286 C72,300 68,312 70,320 L104,320 C106,310 104,298 102,288 L104,284Z" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- right lower leg -->
<path d="M200,286 C208,300 212,312 210,320 L176,320 C174,310 176,298 178,288 L176,284Z" fill="url(#sk8)" stroke="${SKIN_STROKE}" stroke-width="1.2"/>
<!-- buttock contour (visible from front) -->
<path d="M60,90 Q74,108 90,114" stroke="${GUIDE}" stroke-width="0.9" fill="none" opacity="0.4"/>
<path d="M220,90 Q206,108 190,114" stroke="${GUIDE}" stroke-width="0.9" fill="none" opacity="0.4"/>
<!-- liposuction zone markings (dashed) for flanks -->
<path d="M44,10 Q42,50 44,90" stroke="#4A90D9" stroke-width="1.2" fill="none" stroke-dasharray="4,3" opacity="0.35"/>
<path d="M236,10 Q238,50 236,90" stroke="#4A90D9" stroke-width="1.2" fill="none" stroke-dasharray="4,3" opacity="0.35"/>
<text x="30" y="80" font-size="7" fill="${LANDMARK}" opacity="0.6" font-family="Arial,sans-serif">ASIS</text>
<text x="236" y="80" font-size="7" fill="${LANDMARK}" opacity="0.6" font-family="Arial,sans-serif">ASIS</text>
</svg>`

// ── Template data structure ─────────────────────────────────────────────────────────────────────────

const toDataUri = (svg: string) =>
  `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`

interface TemplateView { name: string; svg: string; aspectRatio: number }
interface TemplateGroup { label: string; icon: string; views: TemplateView[] }

const TEMPLATES: Record<string, TemplateGroup> = {
  breast: {
    label: 'Breast', icon: '◎',
    views: [
      { name: 'Front', svg: svgBreastFront, aspectRatio: 280 / 280 },
      { name: 'Front (Body)', svg: svgFemaleBodyFront, aspectRatio: 260 / 500 },
    ],
  },
  body_female: {
    label: 'Body ♀', icon: '⬡',
    views: [
      { name: 'Front', svg: svgFemaleBodyFront, aspectRatio: 260 / 500 },
      { name: 'Back', svg: svgFemaleBodyBack, aspectRatio: 260 / 500 },
    ],
  },
  abdomen: {
    label: 'Abdomen', icon: '◉',
    views: [
      { name: 'Front', svg: svgAbdomenFront, aspectRatio: 280 / 300 },
    ],
  },
  lower_body: {
    label: 'Lower Body', icon: '⬢',
    views: [
      { name: 'Front', svg: svgLowerBodyFront, aspectRatio: 280 / 320 },
    ],
  },
  face_female: {
    label: 'Face ♀', icon: '◑',
    views: [
      { name: 'Front', svg: svgFaceFront, aspectRatio: 240 / 320 },
      { name: 'Lateral', svg: svgFaceLateral, aspectRatio: 240 / 320 },
    ],
  },
  body_male: {
    label: 'Body ♂', icon: '⬡',
    views: [
      { name: 'Front', svg: svgMaleBodyFront, aspectRatio: 260 / 500 },
    ],
  },
}

// ── Types ────────────────────────────────────────────────────────────────────────────────────────────

type Tool = 'select' | 'draw' | 'marker' | 'line' | 'arrow' | 'rect' | 'circle' | 'text' | 'eraser'

interface Stamp { label: string; icon: string; action: (canvas: any, x: number, y: number, color: string) => void }

interface Props {
  encounterId: string
  patientId: string
  orgId: string
  procedureType?: keyof typeof TEMPLATES
  onSave?: (drawingId: string) => void
}

// ── Surgical annotation stamps ────────────────────────────────────────────────────────────────────────

function makeStamps(fabric: any): Stamp[] {
  return [
    {
      label: 'Incision', icon: '—',
      action: (c, x, y, color) => {
        const line = new fabric.Line([x - 40, y, x + 40, y], { stroke: color, strokeWidth: 2, strokeDashArray: [6, 3], selectable: true })
        c.add(line)
      },
    },
    {
      label: 'X Mark', icon: '✕',
      action: (c, x, y, color) => {
        const g = new fabric.Group([
          new fabric.Line([x - 12, y - 12, x + 12, y + 12], { stroke: color, strokeWidth: 2.5 }),
          new fabric.Line([x + 12, y - 12, x - 12, y + 12], { stroke: color, strokeWidth: 2.5 }),
        ], { selectable: true })
        c.add(g)
      },
    },
    {
      label: 'Dot / NAV', icon: '●',
      action: (c, x, y, color) => {
        const circle = new fabric.Circle({ left: x - 6, top: y - 6, radius: 6, fill: color, selectable: true })
        c.add(circle)
      },
    },
    {
      label: 'Hatch Zone', icon: '▩',
      action: (c, x, y, color) => {
        const lines = []
        for (let i = -3; i <= 3; i++) {
          lines.push(new fabric.Line([x - 28 + i * 8, y - 24, x - 28 + i * 8 + 24, y + 24], { stroke: color, strokeWidth: 1.5, opacity: 0.7 }))
        }
        const g = new fabric.Group(lines, { selectable: true })
        c.add(g)
      },
    },
    {
      label: 'IMF Line', icon: '⌒',
      action: (c, x, y, color) => {
        const path = new fabric.Path(`M ${x - 50},${y} Q ${x},${y + 20} ${x + 50},${y}`, {
          stroke: color, strokeWidth: 2, fill: 'transparent', strokeDashArray: [5, 3], selectable: true,
        })
        c.add(path)
      },
    },
    {
      label: 'Circle Zone', icon: '○',
      action: (c, x, y, color) => {
        const circle = new fabric.Circle({ left: x - 30, top: y - 30, radius: 30, fill: 'transparent', stroke: color, strokeWidth: 2, strokeDashArray: [4, 3], selectable: true })
        c.add(circle)
      },
    },
  ]
}

// ── Component ────────────────────────────────────────────────────────────────────────────────────────

export default function SurgicalDrawingTool({ encounterId, patientId, orgId, procedureType = 'breast', onSave }: Props) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  const fabricRef = useRef<any>(null)
  const arrowStartRef = useRef<{ x: number; y: number } | null>(null)
  const penActiveRef = useRef(false)
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const saveDrawingRef = useRef<(() => void) | null>(null)

  const [fabricLoaded, setFabricLoaded] = useState(false)
  const [activeTool, setActiveTool] = useState<Tool>('draw')
  const [activeColor, setActiveColor] = useState('#e74c3c')
  const [brushSize, setBrushSize] = useState(4)
  const [activeTemplate, setActiveTemplate] = useState<keyof typeof TEMPLATES>(
    procedureType in TEMPLATES ? procedureType : 'breast'
  )
  const [activeView, setActiveView] = useState(0)
  const [saving, setSaving] = useState(false)
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saved' | 'error'>('idle')
  const [history, setHistory] = useState<string[]>([])
  const [historyIndex, setHistoryIndex] = useState(-1)
  const [canvasSize, setCanvasSize] = useState({ width: 700, height: 540 })
  const [showStamps, setShowStamps] = useState(false)
  const [isTouch, setIsTouch] = useState(false)
  const [loadingDrawing, setLoadingDrawing] = useState(false)

  // Surgical color palette — includes standard marking colors + skin tones
  const colors = [
    '#1a1a1a', '#e74c3c', '#e67e22', '#f1c40f',
    '#2ecc71', '#3498db', '#9b59b6', '#1abc9c',
    '#0044CC', '#CC0000', '#00AA44', '#8B15CC',
    '#ffffff',
  ]

  // Detect touch device
  useEffect(() => {
    setIsTouch('ontouchstart' in window || navigator.maxTouchPoints > 0)
  }, [])

  // Responsive canvas sizing via ResizeObserver
  useLayoutEffect(() => {
    const el = containerRef.current
    if (!el) return
    const update = (w: number) => {
      const template = TEMPLATES[activeTemplate]
      const ar = template.views[activeView]?.aspectRatio ?? (260 / 500)
      const cw = Math.min(Math.max(w - 32, 320), 1200)
      const ch = Math.round(cw / ar)
      setCanvasSize({ width: cw, height: Math.min(ch, window.innerHeight - 220) })
    }
    const ro = new ResizeObserver(entries => update(entries[0].contentRect.width))
    ro.observe(el)
    update(el.getBoundingClientRect().width)
    return () => ro.disconnect()
  }, [activeTemplate, activeView])

  // Load Fabric.js from CDN (v5 for stability)
  useEffect(() => {
    if ((window as any).fabric) { setFabricLoaded(true); return }
    const script = document.createElement('script')
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/fabric.js/5.3.1/fabric.min.js'
    script.onload = () => setFabricLoaded(true)
    document.head.appendChild(script)
  }, [])

  // Init canvas after Fabric loads
  useEffect(() => {
    if (!fabricLoaded || !canvasRef.current) return
    const fabric = (window as any).fabric

    const canvas = new fabric.Canvas(canvasRef.current, {
      isDrawingMode: true,
      width: canvasSize.width,
      height: canvasSize.height,
      backgroundColor: '#ffffff',
      enableRetinaScaling: true,
      allowTouchScrolling: false,
    })

    canvas.freeDrawingBrush.color = activeColor
    canvas.freeDrawingBrush.width = brushSize
    canvas.freeDrawingBrush.decimate = 2 // Douglas-Peucker simplification on pencil strokes

    canvas.on('object:added', () => pushHistory(canvas))
    canvas.on('object:modified', () => pushHistory(canvas))
    canvas.on('object:removed', () => pushHistory(canvas))

    // Auto-save 2s after last stroke
    canvas.on('path:created', () => {
      if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current)
      autoSaveTimerRef.current = setTimeout(() => {
        setSaveStatus('idle') // will trigger saveDrawing via the effect below
        saveDrawingRef.current?.()
      }, 2000)
    })

    fabricRef.current = canvas
    loadTemplate(canvas, activeTemplate, activeView, canvasSize)

    // Load previously saved drawing for this encounter + template
    const templateKey = `${activeTemplate}:${TEMPLATES[activeTemplate].views[activeView].name}`
    setLoadingDrawing(true)
    supabase.schema('cr').from('surgical_drawings')
      .select('drawing_json')
      .eq('encounter_id', encounterId)
      .eq('template_key', templateKey)
      .maybeSingle()
      .then(({ data }) => {
        if (data?.drawing_json) {
          canvas.loadFromJSON(data.drawing_json, () => {
            canvas.renderAll()
            loadTemplate(canvas, activeTemplate, activeView, canvasSize)
          })
        }
        setLoadingDrawing(false)
      })
      .catch(() => setLoadingDrawing(false))

    return () => {
      if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current)
      canvas.dispose()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fabricLoaded])

  // Resize existing canvas when size changes
  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas || !fabricLoaded) return
    canvas.setWidth(canvasSize.width)
    canvas.setHeight(canvasSize.height)
    canvas.renderAll()
  }, [canvasSize, fabricLoaded])

  // Apple Pencil / pointer pressure sensitivity
  useEffect(() => {
    const el = canvasRef.current
    if (!el || !isTouch) return
    const handlePointerMove = (e: PointerEvent) => {
      if (e.pointerType === 'pen' && fabricRef.current?.isDrawingMode) {
        const pressure = Math.max(e.pressure, 0.2)
        if (fabricRef.current.freeDrawingBrush) {
          fabricRef.current.freeDrawingBrush.width = brushSize * pressure * 2.5
        }
      }
    }
    el.addEventListener('pointermove', handlePointerMove)
    return () => el.removeEventListener('pointermove', handlePointerMove)
  }, [isTouch, brushSize])

  // Prevent page scroll while drawing on iPad
  useEffect(() => {
    const el = canvasRef.current
    if (!el || !isTouch) return
    const prevent = (e: TouchEvent) => { if (fabricRef.current?.isDrawingMode) e.preventDefault() }
    el.addEventListener('touchmove', prevent, { passive: false })
    return () => el.removeEventListener('touchmove', prevent)
  }, [isTouch])

  // Palm rejection — when Apple Pencil is active, block finger touch events from reaching Fabric
  useEffect(() => {
    const el = canvasRef.current
    if (!el || !isTouch) return

    const onPenDown = (e: PointerEvent) => {
      if (e.pointerType === 'pen') penActiveRef.current = true
    }
    const onPenUp = (e: PointerEvent) => {
      if (e.pointerType === 'pen') {
        // Keep pen-active for 120ms to absorb any trailing touch events
        setTimeout(() => { penActiveRef.current = false }, 120)
      }
    }
    const blockPalmTouch = (e: PointerEvent) => {
      if (e.pointerType === 'touch' && penActiveRef.current) e.stopImmediatePropagation()
    }

    el.addEventListener('pointerdown', onPenDown, { capture: true })
    el.addEventListener('pointerup', onPenUp, { capture: true })
    el.addEventListener('pointercancel', onPenUp, { capture: true })
    // Block touch at capture phase so Fabric never sees it
    el.addEventListener('pointerdown', blockPalmTouch, { capture: true })
    el.addEventListener('pointermove', blockPalmTouch, { capture: true })

    return () => {
      el.removeEventListener('pointerdown', onPenDown, { capture: true })
      el.removeEventListener('pointerup', onPenUp, { capture: true })
      el.removeEventListener('pointercancel', onPenUp, { capture: true })
      el.removeEventListener('pointerdown', blockPalmTouch, { capture: true })
      el.removeEventListener('pointermove', blockPalmTouch, { capture: true })
    }
  }, [isTouch])

  // Load SVG template as background image
  const loadTemplate = useCallback((canvas: any, templateKey: string, viewIndex: number, size: { width: number; height: number }) => {
    if (!canvas) return
    const fabric = (window as any).fabric
    const template = TEMPLATES[templateKey]
    if (!template) return
    const view = template.views[viewIndex]
    if (!view) return

    const uri = toDataUri(view.svg)
    canvas.setBackgroundImage(null, () => {})

    fabric.Image.fromURL(uri, (img: any) => {
      if (!img) return
      const scale = Math.min(size.width / (img.width || size.width), size.height / (img.height || size.height)) * 0.88
      img.set({
        scaleX: scale,
        scaleY: scale,
        left: (size.width - (img.width || 0) * scale) / 2,
        top: (size.height - (img.height || 0) * scale) / 2,
        selectable: false,
        evented: false,
        opacity: 1,
      })
      canvas.setBackgroundImage(img, canvas.renderAll.bind(canvas))
    }, { crossOrigin: 'anonymous' })
  }, [])

  // Template / view change
  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas || !fabricLoaded) return
    canvas.clear()
    canvas.backgroundColor = '#ffffff'
    loadTemplate(canvas, activeTemplate, activeView, canvasSize)
    setHistory([])
    setHistoryIndex(-1)
  }, [activeTemplate, activeView, fabricLoaded, loadTemplate, canvasSize])

  // Tool activation
  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas) return
    const fabric = (window as any).fabric
    canvas.isDrawingMode = false
    canvas.selection = false
    canvas.off('mouse:down')

    switch (activeTool) {
      case 'draw':
        canvas.isDrawingMode = true
        canvas.freeDrawingBrush = new fabric.PencilBrush(canvas)
        canvas.freeDrawingBrush.color = activeColor
        canvas.freeDrawingBrush.width = brushSize
        canvas.freeDrawingBrush.decimate = 2
        break
      case 'marker':
        canvas.isDrawingMode = true
        canvas.freeDrawingBrush = new fabric.PencilBrush(canvas)
        canvas.freeDrawingBrush.color = activeColor + '70'
        canvas.freeDrawingBrush.width = brushSize * 5
        canvas.freeDrawingBrush.decimate = 4
        break
      case 'eraser':
        canvas.isDrawingMode = true
        canvas.freeDrawingBrush = new fabric.PencilBrush(canvas)
        canvas.freeDrawingBrush.color = '#ffffff'
        canvas.freeDrawingBrush.width = brushSize * 6
        canvas.freeDrawingBrush.decimate = 4
        break
      case 'select':
        canvas.selection = true
        break
      case 'text':
        canvas.on('mouse:down', handleTextAdd)
        break
      case 'rect':
        canvas.on('mouse:down', handleRectAdd)
        break
      case 'circle':
        canvas.on('mouse:down', handleCircleAdd)
        break
      case 'line':
        canvas.on('mouse:down', handleLineAdd)
        break
      case 'arrow':
        canvas.on('mouse:down', handleArrowStart)
        canvas.on('mouse:up', handleArrowEnd)
        break
    }
    return () => {
      if (canvas) {
        canvas.off('mouse:down')
        canvas.off('mouse:up')
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTool, activeColor, brushSize])

  const handleTextAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    const text = new fabric.IText('Label', {
      left: pointer.x, top: pointer.y,
      fontSize: isTouch ? 18 : 14, fill: activeColor,
      fontFamily: 'Arial, sans-serif', fontWeight: 'bold', editable: true,
    })
    canvas.add(text)
    canvas.setActiveObject(text)
    text.enterEditing()
    canvas.off('mouse:down', handleTextAdd)
    setActiveTool('select')
  }, [activeColor, isTouch])

  const handleRectAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    canvas.add(new fabric.Rect({ left: pointer.x - 40, top: pointer.y - 25, width: 80, height: 50, fill: 'transparent', stroke: activeColor, strokeWidth: 2, rx: 4 }))
    canvas.off('mouse:down', handleRectAdd)
    setActiveTool('select')
  }, [activeColor])

  const handleCircleAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    canvas.add(new fabric.Circle({ left: pointer.x - 30, top: pointer.y - 30, radius: 30, fill: 'transparent', stroke: activeColor, strokeWidth: 2 }))
    canvas.off('mouse:down', handleCircleAdd)
    setActiveTool('select')
  }, [activeColor])

  const handleLineAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    canvas.add(new fabric.Line([pointer.x - 40, pointer.y, pointer.x + 40, pointer.y], { stroke: activeColor, strokeWidth: 2 }))
    canvas.off('mouse:down', handleLineAdd)
    setActiveTool('select')
  }, [activeColor])

  const handleArrowStart = useCallback((opt: any) => {
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    arrowStartRef.current = { x: pointer.x, y: pointer.y }
  }, [])

  const handleArrowEnd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const start = arrowStartRef.current
    if (!start) return
    const pointer = canvas.getPointer(opt.e)
    const dx = pointer.x - start.x
    const dy = pointer.y - start.y
    const len = Math.sqrt(dx * dx + dy * dy)
    if (len < 10) { arrowStartRef.current = null; return }
    const angle = Math.atan2(dy, dx) * 180 / Math.PI
    const headLen = Math.min(20, len * 0.35)
    const line = new fabric.Line([start.x, start.y, pointer.x, pointer.y], { stroke: activeColor, strokeWidth: 2.5 })
    const head = new fabric.Triangle({ width: headLen, height: headLen * 1.2, fill: activeColor, left: pointer.x, top: pointer.y, angle: angle + 90, originX: 'center', originY: 'center' })
    canvas.add(new fabric.Group([line, head], { selectable: true }))
    arrowStartRef.current = null
    setActiveTool('select')
  }, [activeColor])

  const applyStamp = useCallback((stamp: Stamp) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    if (!canvas || !fabric) return
    const cx = canvas.width / 2
    const cy = canvas.height / 2
    stamp.action(canvas, cx, cy, activeColor)
    setShowStamps(false)
    setActiveTool('select')
  }, [activeColor])

  // History
  const pushHistory = (canvas: any) => {
    const json = JSON.stringify(canvas.toJSON())
    setHistory(prev => {
      const next = [...prev.slice(0, historyIndex + 1), json]
      setHistoryIndex(next.length - 1)
      return next
    })
  }

  const undo = () => {
    const canvas = fabricRef.current
    if (!canvas || historyIndex <= 0) return
    const ni = historyIndex - 1
    canvas.loadFromJSON(history[ni], canvas.renderAll.bind(canvas))
    setHistoryIndex(ni)
  }

  const redo = () => {
    const canvas = fabricRef.current
    if (!canvas || historyIndex >= history.length - 1) return
    const ni = historyIndex + 1
    canvas.loadFromJSON(history[ni], canvas.renderAll.bind(canvas))
    setHistoryIndex(ni)
  }

  const clearDrawing = () => {
    const canvas = fabricRef.current
    if (!canvas) return
    canvas.getObjects().forEach((obj: any) => canvas.remove(obj))
    canvas.renderAll()
  }

  const saveDrawing = useCallback(async () => {
    const canvas = fabricRef.current
    if (!canvas) return
    setSaving(true)
    setSaveStatus('idle')
    try {
      const svgData = canvas.toSVG()
      const jsonData = JSON.stringify(canvas.toJSON())
      const templateInfo = `${activeTemplate}:${TEMPLATES[activeTemplate].views[activeView].name}`
      const { data: { user } } = await supabase.auth.getUser()
      const { data, error } = await supabase
        .schema('cr').from('surgical_drawings')
        .upsert(
          {
            encounter_id: encounterId,
            org_id: orgId,
            template_key: templateInfo,
            drawing_svg: svgData,
            drawing_json: jsonData,
            created_by: user?.id ?? null,
            updated_at: new Date().toISOString(),
          },
          { onConflict: 'encounter_id,template_key' },
        )
        .select('id').single()
      if (error) throw error
      setSaveStatus('saved')
      onSave?.(data.id)
      setTimeout(() => setSaveStatus('idle'), 3000)
    } catch { setSaveStatus('error') } finally { setSaving(false) }
  }, [activeTemplate, activeView, encounterId, orgId, onSave])

  // Keep saveDrawingRef current so the auto-save timer closure always calls the latest version
  useEffect(() => { saveDrawingRef.current = saveDrawing }, [saveDrawing])

  const exportPng = () => {
    const canvas = fabricRef.current
    if (!canvas) return
    const a = document.createElement('a')
    a.href = canvas.toDataURL({ format: 'png', multiplier: 2 })
    a.download = `drawing-${encounterId}-${activeTemplate}.png`
    a.click()
  }

  const template = TEMPLATES[activeTemplate]
  const btnH = isTouch ? 48 : 36
  const fontSize = isTouch ? 14 : 12

  const tools: { id: Tool; icon: string; label: string }[] = [
    { id: 'select', icon: '↖', label: 'Select' },
    { id: 'draw', icon: '✏', label: 'Draw' },
    { id: 'marker', icon: '▮', label: 'Marker' },
    { id: 'line', icon: '╱', label: 'Line' },
    { id: 'arrow', icon: '→', label: 'Arrow' },
    { id: 'rect', icon: '▭', label: 'Rect' },
    { id: 'circle', icon: '○', label: 'Circle' },
    { id: 'text', icon: 'T', label: 'Text' },
    { id: 'eraser', icon: '⌫', label: 'Erase' },
  ]

  return (
    <div style={{ fontFamily: 'var(--font-sans, system-ui, sans-serif)', background: '#060e1c', borderRadius: 14, overflow: 'hidden', border: '1px solid rgba(201,169,110,0.2)', display: 'flex', flexDirection: 'column' }}>

      {/* Header */}
      <div style={{ background: '#0a1628', padding: `10px ${isTouch ? 16 : 14}px`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid rgba(201,169,110,0.15)', flexShrink: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
            <polygon points="9,1 17,9 9,17 1,9" fill="none" stroke="#c9a96e" strokeWidth="1.2"/>
            <polygon points="9,5 13,9 9,13 5,9" fill="#c9a96e" opacity="0.4"/>
          </svg>
          <span style={{ color: '#c9a96e', fontSize: 14, fontWeight: 600, letterSpacing: '0.3px' }}>Surgical Drawing</span>
          {isTouch && <span style={{ fontSize: 10, color: 'rgba(201,169,110,0.5)', marginLeft: 4 }}>Apple Pencil ready</span>}
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={exportPng} style={btn('ghost', false, btnH, fontSize)}>Export PNG</button>
          <button onClick={saveDrawing} disabled={saving} style={btn('gold', saving, btnH, fontSize)}>
            {saving ? 'Saving…' : saveStatus === 'saved' ? '✓ Saved' : saveStatus === 'error' ? '✗ Error' : 'Save'}
          </button>
        </div>
      </div>

      {/* Template tabs */}
      <div style={{ background: '#0a1628', borderBottom: '1px solid rgba(255,255,255,0.06)', overflowX: 'auto', flexShrink: 0 }}>
        <div style={{ display: 'flex', gap: 2, padding: `8px ${isTouch ? 12 : 10}px`, minWidth: 'max-content' }}>
          {Object.entries(TEMPLATES).map(([key, tmpl]) => (
            <button key={key}
              onClick={() => { setActiveTemplate(key as keyof typeof TEMPLATES); setActiveView(0) }}
              style={{ padding: `${isTouch ? '10px 16px' : '7px 12px'}`, borderRadius: 8, border: 'none', cursor: 'pointer', fontSize: isTouch ? 14 : 12, fontWeight: 600, whiteSpace: 'nowrap',
                background: activeTemplate === key ? 'rgba(201,169,110,0.2)' : 'rgba(255,255,255,0.04)',
                color: activeTemplate === key ? '#c9a96e' : 'rgba(255,255,255,0.5)',
                outline: activeTemplate === key ? '1px solid rgba(201,169,110,0.4)' : 'none',
              }}>
              <span style={{ marginRight: 6 }}>{tmpl.icon}</span>{tmpl.label}
            </button>
          ))}
        </div>
        {/* View tabs */}
        {template.views.length > 1 && (
          <div style={{ display: 'flex', gap: 2, padding: `0 ${isTouch ? 12 : 10}px 8px`, minWidth: 'max-content' }}>
            {template.views.map((v, i) => (
              <button key={i} onClick={() => setActiveView(i)}
                style={{ padding: `${isTouch ? '8px 14px' : '5px 10px'}`, borderRadius: 6, border: 'none', cursor: 'pointer', fontSize: isTouch ? 13 : 11, fontWeight: 500,
                  background: activeView === i ? 'rgba(0,212,255,0.15)' : 'transparent',
                  color: activeView === i ? '#00d4ff' : 'rgba(255,255,255,0.4)',
                }}>
                {v.name}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Toolbar */}
      <div style={{ background: '#0d1f3c', padding: `${isTouch ? 10 : 8}px ${isTouch ? 12 : 10}px`, display: 'flex', alignItems: 'center', gap: isTouch ? 8 : 6, flexWrap: 'wrap', borderBottom: '1px solid rgba(255,255,255,0.06)', flexShrink: 0 }}>

        {/* Tool buttons */}
        <div style={{ display: 'flex', gap: isTouch ? 6 : 4, flexWrap: 'wrap' }}>
          {tools.map(t => (
            <button key={t.id} title={t.label} onClick={() => { setActiveTool(t.id); setShowStamps(false) }}
              style={{ width: btnH, height: btnH, borderRadius: 8, border: 'none', cursor: 'pointer', fontSize: isTouch ? 16 : 14, fontWeight: 700,
                background: activeTool === t.id ? '#c9a96e' : 'rgba(255,255,255,0.07)',
                color: activeTool === t.id ? '#060e1c' : 'rgba(255,255,255,0.65)',
                boxShadow: activeTool === t.id ? '0 2px 8px rgba(201,169,110,0.4)' : 'none',
                transition: 'all 0.12s',
              }}>
              {t.icon}
            </button>
          ))}
          {/* Stamp button */}
          <div style={{ position: 'relative' }}>
            <button title="Stamps" onClick={() => setShowStamps(s => !s)}
              style={{ width: btnH, height: btnH, borderRadius: 8, border: 'none', cursor: 'pointer', fontSize: isTouch ? 16 : 14, fontWeight: 700,
                background: showStamps ? '#c9a96e' : 'rgba(255,255,255,0.07)',
                color: showStamps ? '#060e1c' : 'rgba(255,255,255,0.65)',
              }}>
              ⊕
            </button>
            {showStamps && fabricLoaded && (
              <div style={{ position: 'absolute', top: btnH + 6, left: 0, background: '#0d1f3c', border: '1px solid rgba(201,169,110,0.3)', borderRadius: 10, padding: 8, zIndex: 100, display: 'flex', flexDirection: 'column', gap: 4, minWidth: 140, boxShadow: '0 8px 32px rgba(0,0,0,0.5)' }}>
                {makeStamps((window as any).fabric).map(stamp => (
                  <button key={stamp.label} onClick={() => applyStamp(stamp)}
                    style={{ display: 'flex', alignItems: 'center', gap: 10, padding: `${isTouch ? 10 : 7}px 12px`, borderRadius: 6, border: 'none', cursor: 'pointer', fontSize: isTouch ? 14 : 12,
                      background: 'rgba(255,255,255,0.05)', color: 'rgba(255,255,255,0.8)', textAlign: 'left',
                    }}>
                    <span style={{ fontSize: 16, color: activeColor }}>{stamp.icon}</span> {stamp.label}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        <div style={{ width: 1, height: 32, background: 'rgba(255,255,255,0.08)', flexShrink: 0 }}/>

        {/* Color palette */}
        <div style={{ display: 'flex', gap: isTouch ? 6 : 4, flexWrap: 'wrap' }}>
          {colors.map(c => (
            <button key={c} onClick={() => setActiveColor(c)}
              style={{ width: isTouch ? 32 : 24, height: isTouch ? 32 : 24, borderRadius: '50%', border: activeColor === c ? '3px solid #c9a96e' : '2px solid transparent',
                background: c, cursor: 'pointer', padding: 0, flexShrink: 0,
                boxShadow: c === '#ffffff' ? 'inset 0 0 0 1px rgba(255,255,255,0.15)' : activeColor === c ? '0 0 8px rgba(201,169,110,0.6)' : 'none',
                transition: 'border 0.1s, box-shadow 0.1s',
              }}/>
          ))}
        </div>

        <div style={{ width: 1, height: 32, background: 'rgba(255,255,255,0.08)', flexShrink: 0 }}/>

        {/* Brush size — use stepper buttons on touch for easier use */}
        {isTouch ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)' }}>Size</span>
            {[2, 4, 8, 14].map(s => (
              <button key={s} onClick={() => setBrushSize(s)}
                style={{ width: 40, height: 40, borderRadius: 8, border: 'none', cursor: 'pointer', fontSize: 11, fontWeight: 600,
                  background: brushSize === s ? 'rgba(201,169,110,0.2)' : 'rgba(255,255,255,0.06)',
                  color: brushSize === s ? '#c9a96e' : 'rgba(255,255,255,0.5)',
                }}>
                {s === 2 ? 'S' : s === 4 ? 'M' : s === 8 ? 'L' : 'XL'}
              </button>
            ))}
          </div>
        ) : (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)' }}>Size</span>
            <input type="range" min={1} max={20} value={brushSize} onChange={e => setBrushSize(Number(e.target.value))}
              style={{ width: 80, accentColor: '#c9a96e' }}/>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.5)', width: 20 }}>{brushSize}</span>
          </div>
        )}

        <div style={{ width: 1, height: 32, background: 'rgba(255,255,255,0.08)', flexShrink: 0 }}/>

        {/* Undo / Redo / Clear */}
        <div style={{ display: 'flex', gap: isTouch ? 6 : 4 }}>
          <button onClick={undo} disabled={historyIndex <= 0} style={btn('ghost', historyIndex <= 0, btnH, fontSize)}>↩</button>
          <button onClick={redo} disabled={historyIndex >= history.length - 1} style={btn('ghost', historyIndex >= history.length - 1, btnH, fontSize)}>↪</button>
          <button onClick={clearDrawing} style={btn('ghost', false, btnH, fontSize)}>✕</button>
        </div>
      </div>

      {/* Canvas */}
      <div ref={containerRef} style={{ background: '#f0f2f5', flex: 1, display: 'flex', justifyContent: 'center', alignItems: 'flex-start', padding: 16, overflowY: 'auto', touchAction: 'pan-y' }}>
        <canvas ref={canvasRef} style={{ borderRadius: 10, boxShadow: '0 4px 28px rgba(0,0,0,0.35)', touchAction: 'none', cursor: activeTool === 'select' ? 'default' : 'crosshair', display: 'block' }}/>
      </div>

      {/* Footer */}
      <div style={{ background: '#060e1c', padding: '7px 14px', borderTop: '1px solid rgba(255,255,255,0.05)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0 }}>
        <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.22)' }}>
          Anatomical references by{' '}
          <a href="https://smart.servier.com" target="_blank" rel="noopener noreferrer" style={{ color: 'rgba(201,169,110,0.45)', textDecoration: 'none' }}>
            Servier Medical Art
          </a>{' '}CC BY 4.0
        </span>
        <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.18)' }}>
          {encounterId.slice(0, 8)}…
        </span>
      </div>
    </div>
  )
}

// ── Style helper ───────────────────────────────────────────────────────────────────────────────────
function btn(variant: 'gold' | 'ghost', disabled: boolean, h = 36, fs = 12): React.CSSProperties {
  return {
    height: h, minWidth: h, paddingLeft: 12, paddingRight: 12,
    borderRadius: 8, border: 'none',
    cursor: disabled ? 'not-allowed' : 'pointer',
    fontSize: fs, fontWeight: 600, opacity: disabled ? 0.35 : 1,
    transition: 'opacity 0.12s',
    background: variant === 'gold' ? '#c9a96e' : 'rgba(255,255,255,0.07)',
    color: variant === 'gold' ? '#060e1c' : 'rgba(255,255,255,0.65)',
  }
}
