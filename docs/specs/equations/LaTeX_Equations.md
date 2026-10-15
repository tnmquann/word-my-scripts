Tôi đã chạy thử và cả hai chiều chuyển đổi đều hoạt động.

**Kết quả chạy thử**
- **Chiều LaTeX sang equation:** chuyển đúng các công thức thử, gồm phân số, Σ/∫/∏, √, `\binom`, ma trận, `cases`, `\overset`, `\underbrace`, `\mathbb`, `\mathcal`, `\mathbf`, `\mathrm`, `\operatorname`, `\lVert…\rVert` và `\text`.
- **Chiều equation sang LaTeX:** chạy trên toàn bộ 8 công thức trong tài liệu và trên các đoạn thử. Kết quả giống LaTeX gốc, chỉ khác vài chi tiết viết tắt như `\le` thành `\leq`.
- **Tài liệu của bạn:** 8 công thức đã chuyển sang LaTeX rồi chuyển ngược lại. Tài liệu vẫn là equation Word, vẫn 9 đoạn, và đoạn thử đã xóa.

**Về lựa chọn**
- **Không có menu trong script:** môi trường này không cho script hiện menu hay nút bấm. Thay vào đó, bạn đổi hai dòng đầu để chọn tính năng và phạm vi.
- **Hai dòng cần đổi:** `FEATURE` nhận `"latex2eq"` hoặc `"eq2latex"`. `SCOPE` nhận `"selection"` (vùng chọn), `"document"` (toàn tài liệu) hoặc `"auto"` (có vùng chọn thì dùng, không thì chuyển cả tài liệu).
- **Bốn nút trong task pane:** nếu bạn tự làm task pane, mỗi nút gọi script này với một trong bốn tổ hợp.

```javascript
// ===== CẤU HÌNH =====
const FEATURE = "latex2eq";   // "latex2eq": LaTeX -> equation | "eq2latex": equation -> LaTeX
const SCOPE = "selection";    // "selection" | "document" | "auto"
const DISPLAY_DELIM = "$$";   // chỉ dùng cho eq2latex: "$$" hoặc "\\[" cho công thức display
// =====================

const FONT="Cambria Math";
const MNS="http://schemas.openxmlformats.org/officeDocument/2006/math";
const WNS="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const esc=s=>s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const RPR=`<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/></w:rPr>`;
function run(t,o={}){
  if(t==="")return "";
  let m="";if(o.nor)m+=`<m:nor m:val="1"/>`;if(o.scr)m+=`<m:scr m:val="${o.scr}"/>`;if(o.sty)m+=`<m:sty m:val="${o.sty}"/>`;
  return `<m:r>${m?`<m:rPr>${m}</m:rPr>`:""}${o.nor?"":RPR}<m:t xml:space="preserve">${esc(t)}</m:t></m:r>`;
}
const sSub=(e,s)=>`<m:sSub><m:e>${e}</m:e><m:sub>${s}</m:sub></m:sSub>`;
const sSup=(e,s)=>`<m:sSup><m:e>${e}</m:e><m:sup>${s}</m:sup></m:sSup>`;
const sSubSup=(e,b,p)=>`<m:sSubSup><m:e>${e}</m:e><m:sub>${b}</m:sub><m:sup>${p}</m:sup></m:sSubSup>`;
const frac=(n,d,t)=>`<m:f>${t?`<m:fPr><m:type m:val="${t}"/></m:fPr>`:""}<m:num>${n}</m:num><m:den>${d}</m:den></m:f>`;
const delim=(e,b,en)=>`<m:d><m:dPr><m:begChr m:val="${b}"/><m:endChr m:val="${en}"/></m:dPr><m:e>${e}</m:e></m:d>`;
const nary=(c,loc,sb,sp,e)=>`<m:nary><m:naryPr><m:chr m:val="${c}"/><m:limLoc m:val="${loc}"/>${sb?"":'<m:subHide m:val="1"/>'}${sp?"":'<m:supHide m:val="1"/>'}</m:naryPr><m:sub>${sb}</m:sub><m:sup>${sp}</m:sup><m:e>${e}</m:e></m:nary>`;
const func=(n,e)=>`<m:func><m:fName>${n}</m:fName><m:e>${e}</m:e></m:func>`;
const limLow=(e,l)=>`<m:limLow><m:e>${e}</m:e><m:lim>${l}</m:lim></m:limLow>`;
const limUpp=(e,l)=>`<m:limUpp><m:e>${e}</m:e><m:lim>${l}</m:lim></m:limUpp>`;
const rad=(d,e)=>`<m:rad><m:radPr>${d?"":'<m:degHide m:val="1"/>'}</m:radPr><m:deg>${d}</m:deg><m:e>${e}</m:e></m:rad>`;
const acc=(c,e)=>`<m:acc><m:accPr><m:chr m:val="${c}"/></m:accPr><m:e>${e}</m:e></m:acc>`;
const bar=(p,e)=>`<m:bar><m:barPr><m:pos m:val="${p}"/></m:barPr><m:e>${e}</m:e></m:bar>`;
const matrix=rows=>`<m:m>${rows.map(r=>`<m:mr>${r.map(c=>`<m:e>${c}</m:e>`).join("")}</m:mr>`).join("")}</m:m>`;
const eqArr=rows=>`<m:eqArr>${rows.map(r=>`<m:e>${r}</m:e>`).join("")}</m:eqArr>`;

const GREEK={alpha:"α",beta:"β",gamma:"γ",delta:"δ",epsilon:"ϵ",varepsilon:"ε",zeta:"ζ",eta:"η",theta:"θ",vartheta:"ϑ",iota:"ι",kappa:"κ",lambda:"λ",mu:"μ",nu:"ν",xi:"ξ",pi:"π",varpi:"ϖ",rho:"ρ",varrho:"ϱ",sigma:"σ",varsigma:"ς",tau:"τ",upsilon:"υ",phi:"ϕ",varphi:"φ",chi:"χ",psi:"ψ",omega:"ω",Gamma:"Γ",Delta:"Δ",Theta:"Θ",Lambda:"Λ",Xi:"Ξ",Pi:"Π",Sigma:"Σ",Upsilon:"Υ",Phi:"Φ",Psi:"Ψ",Omega:"Ω"};
const SYM={times:"×",cdot:"⋅",cdots:"⋯",ldots:"…",dots:"…",vdots:"⋮",ddots:"⋱",pm:"±",mp:"∓",div:"÷",le:"≤",leq:"≤",ge:"≥",geq:"≥",ne:"≠",neq:"≠",approx:"≈",sim:"∼",simeq:"≃",cong:"≅",equiv:"≡",propto:"∝",infty:"∞",to:"→",rightarrow:"→",leftarrow:"←",gets:"←",Rightarrow:"⇒",Leftarrow:"⇐",leftrightarrow:"↔",Leftrightarrow:"⇔",iff:"⇔",implies:"⇒",mapsto:"↦",uparrow:"↑",downarrow:"↓",in:"∈",notin:"∉",ni:"∋",subset:"⊂",subseteq:"⊆",supset:"⊃",supseteq:"⊇",cup:"∪",cap:"∩",emptyset:"∅",varnothing:"∅",forall:"∀",exists:"∃",neg:"¬",lnot:"¬",land:"∧",wedge:"∧",lor:"∨",vee:"∨",partial:"∂",nabla:"∇",circ:"∘",ast:"∗",star:"⋆",bullet:"∙",oplus:"⊕",otimes:"⊗",perp:"⊥",parallel:"∥",angle:"∠",degree:"°",prime:"′",ll:"≪",gg:"≫",setminus:"∖",hbar:"ℏ",ell:"ℓ",Re:"ℜ",Im:"ℑ",aleph:"ℵ",therefore:"∴",because:"∵",mid:"∣",vert:"|",Vert:"‖",lbrace:"{",rbrace:"}",langle:"⟨",rangle:"⟩",colon:":",dagger:"†",ddagger:"‡",triangle:"△",square:"□",Box:"□",checkmark:"✓"};
const NARY={sum:"∑",prod:"∏",coprod:"∐",int:"∫",iint:"∬",iiint:"∭",oint:"∮",bigcup:"⋃",bigcap:"⋂",bigvee:"⋁",bigwedge:"⋀",bigoplus:"⨁",bigotimes:"⨂"};
const FUNCS=["ln","log","lg","exp","sin","cos","tan","cot","sec","csc","arcsin","arccos","arctan","sinh","cosh","tanh","coth","det","dim","ker","deg","hom","arg","mod","gcd","Pr","max","min","sup","inf","lim","limsup","liminf"];
const LIMF=["lim","limsup","liminf","max","min","sup","inf","gcd","Pr","det"];
const FNAME={limsup:"lim sup",liminf:"lim inf"};
const ACC={hat:"̂",widehat:"̂",check:"̌",tilde:"̃",widetilde:"̃",acute:"́",grave:"̀",dot:"̇",ddot:"̈",breve:"̆",bar:"̅",vec:"⃗"};
const PAIRS={lVert:["‖","rVert"],lvert:["|","rvert"],langle:["⟨","rangle"],lfloor:["⌊","rfloor"],lceil:["⌈","rceil"]};
const CLOSE={rVert:"‖",rvert:"|",rangle:"⟩",rfloor:"⌋",rceil:"⌉"};
const TXT={text:{nor:1},textrm:{nor:1},mbox:{nor:1},mathrm:{sty:"p"},textbf:{sty:"b"},mathbf:{sty:"b"},boldsymbol:{sty:"b"},bm:{sty:"b"},mathit:{sty:"i"},textit:{sty:"i"},mathbb:{scr:"double-struck"},mathcal:{scr:"script"},mathscr:{scr:"script"},mathfrak:{scr:"fraktur"},mathsf:{scr:"sans-serif"},mathtt:{scr:"monospace"}};
const ENVD={matrix:["",""],smallmatrix:["",""],pmatrix:["(",")"],bmatrix:["[","]"],Bmatrix:["{","}"],vmatrix:["|","|"],Vmatrix:["‖","‖"]};
const ALPHA_DEF={"double-struck":{cap:0x1D538,low:0x1D552,dig:0x1D7D8,exc:{C:"ℂ",H:"ℍ",N:"ℕ",P:"ℙ",Q:"ℚ",R:"ℝ",Z:"ℤ"},cmd:"\\mathbb"},script:{cap:0x1D49C,low:0x1D4B6,exc:{B:"ℬ",E:"ℰ",F:"ℱ",H:"ℋ",I:"ℐ",L:"ℒ",M:"ℳ",R:"ℛ",e:"ℯ",g:"ℊ",o:"ℴ"},cmd:"\\mathcal"},fraktur:{cap:0x1D504,low:0x1D51E,exc:{C:"ℭ",H:"ℌ",I:"ℑ",R:"ℜ",Z:"ℨ"},cmd:"\\mathfrak"}};
function alphaChar(kind,c){const d=ALPHA_DEF[kind];if(d.exc[c])return d.exc[c];
  if(c>="A"&&c<="Z")return String.fromCodePoint(d.cap+c.charCodeAt(0)-65);
  if(c>="a"&&c<="z")return String.fromCodePoint(d.low+c.charCodeAt(0)-97);
  if(d.dig&&c>="0"&&c<="9")return String.fromCodePoint(d.dig+c.charCodeAt(0)-48);
  return c;}
const REVA={};
for(const kind of Object.keys(ALPHA_DEF)){for(const c of "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"){const u=alphaChar(kind,c);if(u!==c&&!REVA[u])REVA[u]={cmd:ALPHA_DEF[kind].cmd,l:c};}}
const warnings=[];

// ================= LaTeX -> OMML =================
function conv(s){
  let i=0;
  const ws=()=>{while(i<s.length&&/\s/.test(s[i]))i++;};
  function rawArg(){
    ws();
    if(s[i]!=="{"){
      if(s[i]==="\\"){const m=/^\\([A-Za-z]+|[\s\S])/.exec(s.slice(i));i+=m[0].length;return m[0];}
      return i<s.length?s[i++]:"";
    }
    let d=0,st=i+1;
    for(;i<s.length;i++){
      if(s[i]==="\\"){i++;continue;}
      if(s[i]==="{")d++;else if(s[i]==="}"){d--;if(d===0){const r=s.slice(st,i);i++;return r;}}
    }
    return s.slice(st);
  }
  const arg=()=>conv(rawArg());
  function scripts(){
    let sub="",sup="";
    for(;;){
      ws();const c=s[i];
      if(c==="^"){i++;sup+=arg();}
      else if(c==="_"){i++;sub+=arg();}
      else if(c==="'"){i++;sup+=run("′");}
      else break;
    }
    return {sub,sup};
  }
  function attach(b){const {sub,sup}=scripts();if(sub&&sup)return sSubSup(b,sub,sup);if(sub)return sSub(b,sub);if(sup)return sSup(b,sup);return b;}
  const term=()=>attach(atom());
  const STOP=/^(?:[+\-=<>,;]|\\(?:leq?|geq?|neq?|approx|pm|mp|to|rightarrow|Rightarrow|equiv|sim|times|cdot|in)(?![A-Za-z]))/;
  function operand(){let o="";for(;;){ws();if(i>=s.length)break;if(STOP.test(s.slice(i)))break;o+=term();}return o;}
  function funcOperand(){
    ws();if(i>=s.length)return "";
    if(STOP.test(s.slice(i)))return "";
    if(s[i]==="("){
      let d=0;
      for(let j=i;j<s.length;j++){
        if(s[j]==="\\"){j++;continue;}
        if(s[j]==="(")d++;else if(s[j]===")"){d--;if(d===0){const inner=s.slice(i+1,j);i=j+1;return attach(delim(conv(inner),"(",")"));}}
      }
    }
    return term();
  }
  function cleanRaw(r,keepSp){
    r=r.replace(/\\[,;: !]/g," ").replace(/\\([A-Za-z]+)/g,(m,n)=>SYM[n]||GREEK[n]||"").replace(/\\([\s\S])/g,"$1").replace(/[{}]/g,"");
    return keepSp?r:r.replace(/\s+/g,"");
  }
  function delimTok(){
    ws();const m=/^\\([A-Za-z]+|[\s\S])/.exec(s.slice(i));
    if(m){
      i+=m[0].length;const n=m[1];
      if(n===".")return "";
      return ({"{":"{","}":"}","|":"‖",lbrace:"{",rbrace:"}",lVert:"‖",rVert:"‖",Vert:"‖",lvert:"|",rvert:"|",vert:"|",langle:"⟨",rangle:"⟩",lfloor:"⌊",rfloor:"⌋",lceil:"⌈",rceil:"⌉"})[n]??n;
    }
    if(s[i]==="."){i++;return "";}
    return s[i++]||"";
  }
  function findClose(openRe,closeRe){
    let depth=1;const re=new RegExp(`${openRe.source}|${closeRe.source}`,"g");re.lastIndex=i;let m;
    while((m=re.exec(s))){
      if(new RegExp("^"+openRe.source).test(m[0]))depth++;else{depth--;if(depth===0)return m;}
    }
    return null;
  }
  function command(){
    const m=/^\\([A-Za-z]+\*?|[\s\S])/.exec(s.slice(i));i+=m[0].length;
    let n=m[1];const bare=n.replace(/\*$/,"");
    if(n===","||n==="thinspace")return run("\u2009");
    if(n===";"||n===":")return run("\u2005");
    if(n==="!")return "";
    if(n===" ")return run(" ");
    if(n==="quad")return run("\u2003");
    if(n==="qquad")return run("\u2003\u2003");
    if("{}%$&#_".includes(n)&&n.length===1)return run(n);
    if(n==="|")return run("‖");
    if(n==="\\")return run(" ");
    if(n==="frac"||n==="dfrac"||n==="tfrac"||n==="cfrac"){const a=arg(),b=arg();return frac(a,b);}
    if(n==="binom"||n==="dbinom"||n==="tbinom"){const a=arg(),b=arg();return delim(frac(a,b,"noBar"),"(",")");}
    if(n==="sqrt"){ws();let d="";if(s[i]==="["){const j=s.indexOf("]",i);d=conv(s.slice(i+1,j));i=j+1;}return rad(d,arg());}
    if(bare==="operatorname"||n==="mathop"){
      const name=cleanRaw(rawArg());
      if(n==="mathop")return run(name,{sty:"p"});
      return func(attach(run(name,{sty:"p"})),funcOperand());
    }
    if(TXT[n]){const o=TXT[n];let t=cleanRaw(rawArg(),!!o.nor);if(o.scr&&ALPHA_DEF[o.scr])return run([...t].map(c=>alphaChar(o.scr,c)).join(""));return run(t,o);}
    if(FUNCS.includes(n)){
      const nm=FNAME[n]||n;ws();
      if(LIMF.includes(n)&&s[i]==="_"){i++;const lim=arg();return func(limLow(run(nm,{sty:"p"}),lim),funcOperand());}
      return func(attach(run(nm,{sty:"p"})),funcOperand());
    }
    if(NARY[n]){
      const rest=s.slice(i);const lm=/^\s*\\(no)?limits(?![A-Za-z])/.exec(rest);let lo=null;if(lm){i+=lm[0].length;lo=lm[1]?"subSup":"undOvr";}
      const sc=scripts();
      const loc=lo||((n.includes("int"))?"subSup":"undOvr");
      return nary(NARY[n],loc,sc.sub,sc.sup,operand());
    }
    if(ACC[n])return acc(ACC[n],arg());
    if(n==="overline")return bar("top",arg());
    if(n==="underline")return bar("bot",arg());
    if(n==="overset"||n==="stackrel"){const a=arg(),b=arg();return limUpp(b,a);}
    if(n==="underset"){const a=arg(),b=arg();return limLow(b,a);}
    if(n==="underbrace"||n==="overbrace"){
      const e=arg();const top=n==="overbrace";
      const g="<m:groupChr><m:groupChrPr><m:chr m:val=\""+(top?"⏞":"⏟")+"\"/><m:pos m:val=\""+(top?"top":"bot")+"\"/><m:vertJc m:val=\""+(top?"bot":"top")+"\"/></m:groupChrPr><m:e>"+e+"</m:e></m:groupChr>";
      ws();if(!top&&s[i]==="_"){i++;return limLow(g,arg());}if(top&&s[i]==="^"){i++;return limUpp(g,arg());}return g;
    }
    if(n==="left"){
      const open=delimTok();const mm=findClose(/\\left(?![A-Za-z])/,/\\right(?![A-Za-z])/);
      const inner=s.slice(i,mm?mm.index:s.length);i=mm?mm.index+mm[0].length:s.length;
      const close=mm?delimTok():"";
      return delim(conv(inner),open,close);
    }
    if(n==="middle")return run(delimTok());
    if(PAIRS[n]){
      const [ch,cl]=PAIRS[n];
      const mm=findClose(new RegExp("\\\\"+n+"(?![A-Za-z])"),new RegExp("\\\\"+cl+"(?![A-Za-z])"));
      if(mm){const inner=s.slice(i,mm.index);i=mm.index+mm[0].length;return delim(conv(inner),ch,CLOSE[cl]);}
      return run(ch);
    }
    if(CLOSE[n])return run(CLOSE[n]);
    if(n==="begin"){
      const env=rawArg(),end="\\end{"+env+"}";let j=s.indexOf(end,i);
      const body=s.slice(i,j<0?s.length:j);i=j<0?s.length:j+end.length;
      let rows=body.split(/\\\\/).map(r=>r.trim()).filter((r,k,a)=>r!==""||k<a.length-1);
      const e=env.replace(/\*$/,"");
      if(ENVD[e]){const [b,c]=ENVD[e];const mx=matrix(rows.map(r=>r.split("&").map(c=>conv(c))));return b||c?delim(mx,b,c):mx;}
      if(e==="cases")return delim(eqArr(rows.map(r=>conv(r.replace(/&/g," \\quad ")))),"{","");
      return eqArr(rows.map(r=>conv(r)));
    }
    if(["displaystyle","textstyle","scriptstyle","scriptscriptstyle","limits","nolimits","big","Big","bigg","Bigg","nonumber","notag","hline"].includes(n))return "";
    if(n==="label"||n==="tag"||n==="phantom"||n==="hspace"||n==="vspace"){rawArg();return "";}
    if(GREEK[n])return run(GREEK[n]);
    if(SYM[n])return run(SYM[n]);
    warnings.push("\\"+n);
    return run(n,{sty:"p"});
  }
  function atom(){
    ws();const c=s[i];
    if(c==="{")return conv(rawArg());
    if(c==="\\")return command();
    if(/[0-9]/.test(c)){const m=/^[0-9]+(?:\.[0-9]+)?/.exec(s.slice(i));i+=m[0].length;return run(m[0]);}
    const ch=String.fromCodePoint(s.codePointAt(i));i+=ch.length;
    if(ch==="-")return run("−");
    if(ch==="*")return run("∗");
    if(ch==="~")return run("\u00A0");
    return run(ch);
  }
  let out="";for(;;){ws();if(i>=s.length)break;out+=term();}
  return out;
}
const wrapPkg=(x,display)=>`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p>${display?`<m:oMathPara><m:oMath>${x}</m:oMath></m:oMathPara>`:`<m:oMath>${x}</m:oMath>`}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

// ================= OMML -> LaTeX =================
const REV={"×":"\\times","⋅":"\\cdot","⋯":"\\cdots","…":"\\ldots","⋮":"\\vdots","⋱":"\\ddots","±":"\\pm","∓":"\\mp","÷":"\\div","≤":"\\leq","≥":"\\geq","≠":"\\neq","≈":"\\approx","∼":"\\sim","≃":"\\simeq","≅":"\\cong","≡":"\\equiv","∝":"\\propto","∞":"\\infty","→":"\\to","←":"\\leftarrow","⇒":"\\Rightarrow","⇐":"\\Leftarrow","↔":"\\leftrightarrow","⇔":"\\Leftrightarrow","↦":"\\mapsto","↑":"\\uparrow","↓":"\\downarrow","∈":"\\in","∉":"\\notin","∋":"\\ni","⊂":"\\subset","⊆":"\\subseteq","⊃":"\\supset","⊇":"\\supseteq","∪":"\\cup","∩":"\\cap","∅":"\\emptyset","∀":"\\forall","∃":"\\exists","¬":"\\neg","∧":"\\wedge","∨":"\\vee","∂":"\\partial","∇":"\\nabla","∘":"\\circ","⋆":"\\star","∙":"\\bullet","⊕":"\\oplus","⊗":"\\otimes","⊥":"\\perp","∥":"\\parallel","∠":"\\angle","°":"^{\\circ}","≪":"\\ll","≫":"\\gg","∖":"\\setminus","ℏ":"\\hbar","ℓ":"\\ell","ℜ":"\\Re","ℑ":"\\Im","ℵ":"\\aleph","∴":"\\therefore","∵":"\\because","∣":"\\mid","⟨":"\\langle","⟩":"\\rangle","†":"\\dagger","‡":"\\ddagger","△":"\\triangle","□":"\\square","✓":"\\checkmark","‖":"\\|"};
for(const [k,v] of Object.entries(GREEK))REV[v]="\\"+k;
const NREV={};for(const [k,v] of Object.entries(NARY))NREV[v]="\\"+k;
const AREV={"̂":"\\hat","̃":"\\tilde","̇":"\\dot","̈":"\\ddot","̌":"\\check","́":"\\acute","̀":"\\grave","̆":"\\breve","̅":"\\bar","̄":"\\bar","⃗":"\\vec"};
const SCRCMD={"double-struck":"\\mathbb",script:"\\mathcal",fraktur:"\\mathfrak","sans-serif":"\\mathsf",monospace:"\\mathtt"};
const L=n=>n.localName;
const els=n=>n?[...n.childNodes].filter(c=>c.nodeType===1):[];
const mk=(n,name)=>els(n).find(c=>c.namespaceURI===MNS&&L(c)===name);
const val=x=>x?(x.hasAttributeNS(MNS,"val")?x.getAttributeNS(MNS,"val"):""):null;
let inEq=false,inFName=false;
function mapChars(t){
  let o="";
  for(const c of t){
    const cp=c.codePointAt(0);
    if(REVA[c]&&!REV[c]){o+=REVA[c].cmd+"{"+REVA[c].l+"}";continue;}
    let x=c;if(cp>=0x1D400&&cp<=0x1D7FF)x=c.normalize("NFKC");
    if(x==="\u2009")o+="\\,";
    else if(x==="\u2005")o+="\\;";
    else if(x==="\u2003")o+="\\quad ";
    else if(x==="\u00A0")o+=" ";
    else if(x==="−")o+="-";
    else if(x==="∗")o+="*";
    else if(x==="′")o+="'";
    else if(REV[x])o+=REV[x]+" ";
    else if("{}%$#_".includes(x))o+="\\"+x;
    else if(x==="&")o+=inEq?"&":"\\&";
    else if(x==="\\")o+="\\backslash ";
    else o+=x;
  }
  return o;
}
const fnCmd=t=>{t=t.trim();const k=t.replace(/\s+/g,"");if(FUNCS.includes(k))return "\\"+k+" ";return "\\operatorname{"+t+"}";};
function runToLatex(r){
  const t=els(r).filter(c=>L(c)==="t").map(c=>c.textContent).join("");
  if(t==="")return "";
  const rp=mk(r,"rPr");let nor=false,sty=null,scr=null;
  if(rp){const nn=mk(rp,"nor");nor=!!nn&&val(nn)!=="0";sty=val(mk(rp,"sty"));scr=val(mk(rp,"scr"));}
  if(nor)return "\\text{"+t.replace(/([{}%$#_&])/g,"\\$1")+"}";
  if(inFName&&/^[A-Za-z]{2,}(\s?[A-Za-z]+)?$/.test(t.trim()))return fnCmd(t);
  const wrapCmd=scr&&SCRCMD[scr]?SCRCMD[scr]:(sty==="p"?"\\mathrm":sty==="b"?"\\mathbf":sty==="bi"?"\\boldsymbol":null);
  if(wrapCmd)return t.split(/([A-Za-z]+)/).map((seg,k)=>k%2?wrapCmd+"{"+seg+"}":mapChars(seg)).join("");
  return mapChars(t);
}
const isPr=n=>/Pr$/.test(L(n));
const kids=n=>els(n).filter(c=>c.namespaceURI===MNS&&!isPr(c));
const ch=n=>n?els(n).map(cv).join(""):"";
const tok=s=>/^(\\[A-Za-z]+|\\.|[^\\{}\s])$/.test(s.trim());
const sc=s=>{s=s.trim();return tok(s)?s:"{"+s+"}";};
const baseOf=eNode=>{
  const s=ch(eNode).trim();const k=kids(eNode);
  if(tok(s)||/^[0-9]+(\.[0-9]+)?$/.test(s))return s;
  if(k.length===1&&["d","f","func","acc","bar","rad"].includes(L(k[0])))return s;
  return "{"+s+"}";
};
const needBrace=s=>{let d=0;for(let i=0;i<s.length;i++){const c=s[i];if(c==="\\"){i++;continue;}if("([{".includes(c))d++;else if(")]}".includes(c))d--;else if(d===0&&"+-=<>,".includes(c))return true;}return false;};
const tall=n=>["f","nary","rad","m","eqArr","limLow","limUpp"].some(t=>n.getElementsByTagNameNS(MNS,t).length>0);
function dl(c){if(c==="")return ".";if(c==="{")return "\\{";if(c==="}")return "\\}";if(c==="‖")return "\\|";if(REV[c])return REV[c]+" ";return c;}
function matrixLatex(m,env){
  const rows=els(m).filter(c=>L(c)==="mr").map(r=>els(r).filter(c=>L(c)==="e").map(e=>ch(e).trim()).join(" & "));
  return "\\begin{"+env+"} "+rows.join(" \\\\ ")+" \\end{"+env+"}";
}
function eqArrRows(n){const old=inEq;inEq=true;const rows=els(n).filter(c=>L(c)==="e").map(e=>ch(e).trim());inEq=old;return rows;}
function cv(n){
  if(n.nodeType!==1||n.namespaceURI!==MNS||isPr(n))return "";
  switch(L(n)){
    case "r":return runToLatex(n);
    case "f":{const num=ch(mk(n,"num")).trim(),den=ch(mk(n,"den")).trim();const ty=val(mk(mk(n,"fPr"),"type"));
      if(ty==="noBar")return "\\binom{"+num+"}{"+den+"}";
      if(ty==="lin"||ty==="skw")return "{"+num+"}/{"+den+"}";
      return "\\frac{"+num+"}{"+den+"}";}
    case "sSub":return baseOf(mk(n,"e"))+"_"+sc(ch(mk(n,"sub")));
    case "sSup":return baseOf(mk(n,"e"))+"^"+sc(ch(mk(n,"sup")));
    case "sSubSup":return baseOf(mk(n,"e"))+"_"+sc(ch(mk(n,"sub")))+"^"+sc(ch(mk(n,"sup")));
    case "sPre":return "{}_"+sc(ch(mk(n,"sub")))+"^"+sc(ch(mk(n,"sup")))+baseOf(mk(n,"e"));
    case "nary":{
      const cn=mk(mk(n,"naryPr"),"chr");const chr=cn?val(cn):"∫";
      const sub=ch(mk(n,"sub")).trim(),sup=ch(mk(n,"sup")).trim();let e=ch(mk(n,"e")).trim();
      if(needBrace(e))e="{"+e+"}";
      return (NREV[chr]||"\\int")+(sub?"_"+sc(sub):"")+(sup?"^"+sc(sup):"")+" "+e;}
    case "func":{
      const fn=mk(n,"fName");const first=els(fn)[0];let head;
      if(first&&L(first)==="limLow"){
        const old=inFName;inFName=true;const b=ch(mk(first,"e")).trim();inFName=old;
        head=b+"_"+sc(ch(mk(first,"lim")));
      }else{const old=inFName;inFName=true;head=ch(fn).trim();inFName=old;}
      let e=ch(mk(n,"e")).trim();if(needBrace(e))e="{"+e+"}";
      return head+(/^(\(|\\left)/.test(e)&&head.endsWith("}")?"":" ")+e;}
    case "d":{
      const bn=mk(mk(n,"dPr"),"begChr"),en=mk(mk(n,"dPr"),"endChr"),sn=mk(mk(n,"dPr"),"sepChr");
      const beg=bn?val(bn):"(",end=en?val(en):")",sep=sn?val(sn):"|";
      const es=els(n).filter(c=>L(c)==="e");
      if(es.length===1){
        const k=kids(es[0]);
        if(k.length===1&&L(k[0])==="f"&&val(mk(mk(k[0],"fPr"),"type"))==="noBar"&&beg==="("&&end===")")return "\\binom{"+ch(mk(k[0],"num")).trim()+"}{"+ch(mk(k[0],"den")).trim()+"}";
        if(k.length===1&&L(k[0])==="m"){const env={"(":"pmatrix","[":"bmatrix","{":"Bmatrix","|":"vmatrix","‖":"Vmatrix"}[beg];if(env&&({"(":")","[":"]","{":"}","|":"|","‖":"‖"})[beg]===end)return matrixLatex(k[0],env);}
        if(k.length===1&&L(k[0])==="eqArr"&&beg==="{"&&end==="")return "\\begin{cases} "+eqArrRows(k[0]).map(r=>r.replace("\\quad ","& ")).join(" \\\\ ")+" \\end{cases}";
      }
      const inner=es.map(e=>ch(e).trim()).join(" "+(sep==="|"?"\\mid ":dl(sep))+" ");
      if(beg==="‖"&&end==="‖")return "\\lVert "+inner+" \\rVert";
      if(beg==="|"&&end==="|")return "\\lvert "+inner+" \\rvert";
      if(beg==="⟨"&&end==="⟩")return "\\langle "+inner+" \\rangle";
      if(beg==="⌊"&&end==="⌋")return "\\lfloor "+inner+" \\rfloor";
      if(beg==="⌈"&&end==="⌉")return "\\lceil "+inner+" \\rceil";
      if(!tall(n)){
        if(beg==="("&&end===")")return "("+inner+")";
        if(beg==="["&&end==="]")return "["+inner+"]";
        if(beg==="{"&&end==="}")return "\\{"+inner+"\\}";
      }
      return "\\left"+dl(beg)+inner+"\\right"+dl(end);}
    case "rad":{const dg=ch(mk(n,"deg")).trim();return "\\sqrt"+(dg?"["+dg+"]":"")+"{"+ch(mk(n,"e")).trim()+"}";}
    case "acc":{const cn=mk(mk(n,"accPr"),"chr");const c=cn?val(cn):"̂";return (AREV[c]||"\\hat")+"{"+ch(mk(n,"e")).trim()+"}";}
    case "bar":{const p=val(mk(mk(n,"barPr"),"pos"));return (p==="bot"?"\\underline{":"\\overline{")+ch(mk(n,"e")).trim()+"}";}
    case "limLow":{const b=ch(mk(n,"e")).trim(),l=ch(mk(n,"lim")).trim();
      if(kids(mk(n,"e")).length===1&&L(kids(mk(n,"e"))[0])==="groupChr")return b+"_"+sc(l);
      if(LIMF.includes(mk(n,"e").textContent.trim()))return fnCmd(mk(n,"e").textContent)+"_"+sc(l);
      return "\\underset{"+l+"}{"+b+"}";}
    case "limUpp":{const b=ch(mk(n,"e")).trim(),l=ch(mk(n,"lim")).trim();
      if(kids(mk(n,"e")).length===1&&L(kids(mk(n,"e"))[0])==="groupChr")return b+"^"+sc(l);
      return "\\overset{"+l+"}{"+b+"}";}
    case "groupChr":{const c=val(mk(mk(n,"groupChrPr"),"chr"));const p=val(mk(mk(n,"groupChrPr"),"pos"));
      return (c==="⏞"||p==="top"?"\\overbrace{":"\\underbrace{")+ch(mk(n,"e")).trim()+"}";}
    case "m":return matrixLatex(n,"matrix");
    case "eqArr":return "\\begin{aligned} "+eqArrRows(n).join(" \\\\ ")+" \\end{aligned}";
    default:return ch(n);
  }
}
const clean=s=>s.replace(/(\\[A-Za-z]+) +(?=[^A-Za-z\s\\])/g,"$1").replace(/\s{2,}/g," ").trim();
const omathLatex=m=>clean(ch(m));
function ommlToLatexXml(xml){
  const doc=new DOMParser().parseFromString(xml,"application/xml");
  if(doc.getElementsByTagName("parsererror").length)return null;
  const mk2=(t)=>[...doc.getElementsByTagNameNS(MNS,t)];
  const targets=[...mk2("oMathPara"),...mk2("oMath").filter(m=>!(m.parentNode&&m.parentNode.namespaceURI===MNS&&L(m.parentNode)==="oMathPara"))];
  if(!targets.length)return null;
  const disp=DISPLAY_DELIM==="\\["?["\\[","\\]"]:["$$","$$"];
  for(const t of targets){
    const isPara=L(t)==="oMathPara";
    const latex=isPara?els(t).filter(c=>L(c)==="oMath").map(omathLatex).join(" "):omathLatex(t);
    const text=isPara?disp[0]+latex+disp[1]:"$"+latex+"$";
    const r=doc.createElementNS(WNS,"w:r");
    const rp0=t.getElementsByTagNameNS(WNS,"rPr")[0];
    if(rp0){const rp=rp0.cloneNode(true);[...rp.getElementsByTagNameNS(WNS,"rFonts")].forEach(x=>x.parentNode.removeChild(x));if(rp.childNodes.length)r.appendChild(rp);}
    const tt=doc.createElementNS(WNS,"w:t");tt.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");tt.textContent=text;r.appendChild(tt);
    t.parentNode.replaceChild(r,t);
  }
  return {xml:new XMLSerializer().serializeToString(doc),count:targets.length};
}

// ================= Drivers =================
const sel=context.document.getSelection();sel.load("isEmpty");await context.sync();
const useSel=SCOPE==="selection"||(SCOPE==="auto"&&!sel.isEmpty);
const result={feature:FEATURE,scope:useSel?"selection":"document",converted:0,skipped:[]};

async function latex2eq(){
  // $$...$$ | \[...\] | $...$   (tiền tệ kiểu "$5 and $6" được bỏ qua)
  const RE=/\$\$([\s\S]+?)\$\$|\\\[([\s\S]+?)\\\]|(?<![\\$])\$(?![\s$])((?:\\[\s\S]|[^$\\\n])+?)(?<![\s\\])\$(?![\d$])/g;
  const countBefore=(t,nd,pos)=>{let c=0,k=0,j;while((j=t.indexOf(nd,k))!==-1&&j<pos){c++;k=j+nd.length;}return c;};
  const q=s=>s.replace(/\^/g,"^^");              // ^ là ký tự đặc biệt của Word Search
  const paras=useSel?sel.paragraphs:context.document.body.paragraphs;
  paras.load("items/text");await context.sync();
  for(const p of paras.items){
    const text=p.text;const ms=[...text.matchAll(RE)];
    for(const m of ms.reverse()){                  // thay từ cuối lên đầu để offset không lệch
      const raw=m[0],start=m.index,end=start+raw.length;
      const latex=(m[1]??m[2]??m[3]).trim(),display=m[1]!==undefined||m[2]!==undefined;
      try{
        let r;
        if(raw.length<=250){
          const h=p.search(q(raw),{matchCase:true});h.load("items");await context.sync();
          r=h.items[countBefore(text,raw,start)];
        }else{                                     // Word Search giới hạn 255 ký tự: tìm đầu + cuối
          const head=raw.slice(0,100),tail=raw.slice(-100);
          const h=p.search(q(head),{matchCase:true}),t=p.search(q(tail),{matchCase:true});
          h.load("items");t.load("items");await context.sync();
          const a=h.items[countBefore(text,head,start)],b=t.items[countBefore(text,tail,end-tail.length)];
          if(a&&b)r=a.getRange("Start").expandTo(b.getRange("End"));
        }
        if(!r){result.skipped.push(raw.slice(0,40));continue;}
        if(useSel){const c=r.compareLocationWith(sel);await context.sync();
          if(!["Inside","InsideStart","InsideEnd","Equal"].includes(c.value))continue;}
        r.insertOoxml(wrapPkg(conv(latex),display),"Replace");
        await context.sync();result.converted++;
      }catch(e){result.skipped.push(raw.slice(0,40)+" ("+e.code+")");}
    }
  }
  result.unknownCommands=[...new Set(warnings)];
}
async function eq2latex(){
  const unsafe=/<w:drawing|<w:pict|<w:fldChar|<w:fldSimple|<w:object/;
  if(useSel){
    const ox=sel.getOoxml();await context.sync();
    const out=ommlToLatexXml(ox.value);
    if(out){sel.insertOoxml(out.xml,"Replace");await context.sync();result.converted+=out.count;}
    return;
  }
  const paras=context.document.body.paragraphs;paras.load("items");await context.sync();
  const oxs=paras.items.map(p=>p.getOoxml());await context.sync();
  for(let k=paras.items.length-1;k>=0;k--){
    const xml=oxs[k].value;if(!xml.includes("<m:oMath"))continue;
    if(unsafe.test(xml)){result.skipped.push("paragraph "+k+" (contains field/image)");continue;}
    const out=ommlToLatexXml(xml);if(!out)continue;
    paras.items[k].insertOoxml(out.xml,"Replace");await context.sync();result.converted+=out.count;
  }
}
if(FEATURE==="latex2eq")await latex2eq();
else if(FEATURE==="eq2latex")await eq2latex();
else throw new Error("FEATURE phải là latex2eq hoặc eq2latex");
return result;
```

**Lưu ý khi dùng**
- **Chuyển ngược về LaTeX:** `eq2latex` ở chế độ `"document"` bỏ qua đoạn có field hoặc ảnh và ghi vào `skipped`. Lý do là script phải thay lại cả đoạn. Chế độ `"selection"` không có giới hạn này.
- **Công thức đứng một mình trong đoạn:** Word coi đó là display, nên khi chuyển ngược sẽ ra `$$…$$` dù lúc đầu bạn viết `$…$`.
- **Dấu cách trong công thức:** các khoảng trắng quanh `=` và `+` do Word tự thêm khi dựng equation, nên LaTeX xuất ra có thể có dấu cách khác bản gốc.
- **Lệnh chưa hỗ trợ:** `latex2eq` ghi các lệnh LaTeX chưa hỗ trợ vào `unknownCommands` của kết quả trả về. `eq2latex` chỉ dịch đúng các cấu trúc phổ biến. Cấu trúc hiếm như `m:borderBox` sẽ chỉ lấy phần nội dung.
