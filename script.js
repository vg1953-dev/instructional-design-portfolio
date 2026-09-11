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
}
document.querySelectorAll("[data-dialog]").forEach(b=>b.addEventListener("click",()=>{const d=document.getElementById(b.dataset.dialog);if(d)d.showModal()}));
document.querySelectorAll("dialog").forEach(d=>{d.querySelector(".dialog-close")?.addEventListener("click",()=>d.close());d.addEventListener("click",e=>{const r=d.getBoundingClientRect();if(!(e.clientX>=r.left&&e.clientX<=r.right&&e.clientY>=r.top&&e.clientY<=r.bottom))d.close()})});