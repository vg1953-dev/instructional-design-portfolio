document.getElementById("year").textContent=new Date().getFullYear();
const experienceSection=document.querySelector(".experience");
if(experienceSection){
  experienceSection.id="experience";
  const nav=document.querySelector(".nav");
  const contactLink=nav?.querySelector('a[href="#contact"]');
  if(nav&&!nav.querySelector('a[href="#experience"]')){
    const experienceLink=document.createElement("a");
    experienceLink.href="#experience";
    experienceLink.textContent="Experience";
    nav.insertBefore(experienceLink,contactLink||null);
  }
  if(nav&&!nav.querySelector('a[href="finance.html"]')){
    const financeLink=document.createElement("a");
    financeLink.href="finance.html";
    financeLink.textContent="Finance";
    nav.insertBefore(financeLink,contactLink||null);
  }
}
const caseGrid=document.querySelector(".case-grid");
if(caseGrid&&!caseGrid.querySelector('[data-finance-case="true"]')){
  const financeCard=document.createElement("article");
  financeCard.className="case-card featured";
  financeCard.dataset.financeCase="true";
  financeCard.innerHTML='<div class="case-number">05</div><p class="case-tag">Finance & Analytics • Excel Modeling • Dashboards</p><h3>Manufacturing Cost & Profitability Analysis</h3><p class="case-summary">Built a driver-based manufacturing finance model covering unit economics, budget-versus-actual variance, break-even analysis, scenarios, sensitivities, and executive dashboards.</p><a class="text-button" href="finance.html">View finance case study →</a>';
  caseGrid.appendChild(financeCard);
}
document.querySelectorAll("[data-dialog]").forEach(b=>b.addEventListener("click",()=>{const d=document.getElementById(b.dataset.dialog);if(d)d.showModal()}));
document.querySelectorAll("dialog").forEach(d=>{d.querySelector(".dialog-close")?.addEventListener("click",()=>d.close());d.addEventListener("click",e=>{const r=d.getBoundingClientRect();if(!(e.clientX>=r.left&&e.clientX<=r.right&&e.clientY>=r.top&&e.clientY<=r.bottom))d.close()})});