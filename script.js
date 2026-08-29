const header=document.querySelector('#header');
const menu=document.querySelector('#menu');
const nav=document.querySelector('#nav');
const lightbox=document.querySelector('#lightbox');
const lightboxImage=lightbox.querySelector('img');
const lightboxLabel=lightbox.querySelector('p');

const updateHeader=()=>header.classList.toggle('solid',window.scrollY>40);
updateHeader();
window.addEventListener('scroll',updateHeader,{passive:true});

menu.addEventListener('click',()=>{
  const open=!nav.classList.contains('open');
  nav.classList.toggle('open',open);
  menu.classList.toggle('open',open);
  menu.setAttribute('aria-expanded',String(open));
  document.body.classList.toggle('lock',open);
});
nav.querySelectorAll('a').forEach(link=>link.addEventListener('click',()=>{
  nav.classList.remove('open');menu.classList.remove('open');menu.setAttribute('aria-expanded','false');document.body.classList.remove('lock');
}));

document.querySelectorAll('img[data-fallback]').forEach(image=>image.addEventListener('error',()=>{
  if(image.src!==image.dataset.fallback) image.src=image.dataset.fallback;
},{once:true}));

document.querySelectorAll('.photo').forEach(button=>button.addEventListener('click',()=>{
  const image=button.querySelector('img');
  lightboxImage.src=image.currentSrc||image.src;
  lightboxImage.alt=image.alt;
  lightboxLabel.textContent=button.dataset.label||'';
  lightbox.hidden=false;
  document.body.classList.add('lock');
  lightbox.querySelector('button').focus();
}));
const closeLightbox=()=>{lightbox.hidden=true;document.body.classList.remove('lock')};
lightbox.querySelector('button').addEventListener('click',closeLightbox);
lightbox.addEventListener('click',event=>{if(event.target===lightbox) closeLightbox()});
document.addEventListener('keydown',event=>{if(event.key==='Escape'&&!lightbox.hidden) closeLightbox()});

const bookingFrame=document.querySelector('#booking-frame');
const bookingFrameWrap=document.querySelector('#booking-frame-wrap');
if(bookingFrame&&bookingFrameWrap){
  const bookingUrl=window.NV11_BOOKING_APP_URL||'';
  if(bookingUrl&&/^https:\/\/script\.google\.com\//.test(bookingUrl)){
    bookingFrame.src=bookingUrl;
    bookingFrame.addEventListener('load',()=>bookingFrameWrap.classList.add('loaded'));
  }else{
    const loading=bookingFrameWrap.querySelector('#booking-loading');
    if(loading){
      loading.querySelector('strong').textContent='予約フォームは現在準備中です';
      loading.querySelector('span').textContent='公開設定が完了するまで、今しばらくお待ちください。';
    }
  }
  window.addEventListener('message',event=>{
    if(!event.data||event.data.type!=='nv11-booking-height')return;
    const height=Number(event.data.height);
    if(Number.isFinite(height)&&height>=600&&height<=1800)bookingFrame.style.height=`${height}px`;
  });
}
