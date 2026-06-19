const projects = [
  { type: 'A.01 / Essential', title: 'Точный дом<br>для нового ритма.', description: 'Лаконичная базовая конфигурация с панорамным жилым контуром и полной инженерной автономией.', format: 'Компактный', floors: '1–2 этажа', plan: 'Свободная' },
  { type: 'A.02 / Family', title: 'Пространство<br>для большой семьи.', description: 'Расширенный жилой сценарий: приватные зоны, общее светлое ядро и возможность менять комнаты со временем.', format: 'Семейный', floors: '2 этажа', plan: '4–6 комнат' },
  { type: 'A.03 / Courtyard', title: 'Архитектура,<br>обращённая внутрь.', description: 'Дом формируется вокруг защищённого двора — места для тишины, сада, встреч и продолжения интерьера наружу.', format: 'С внутренним двором', floors: '1–2 этажа', plan: 'Секционная' },
  { type: 'A.Custom / Bespoke', title: 'Ваш сценарий.<br>Наша система.', description: 'Индивидуальная архитектура на базе проверенной технологии A-House — для особого участка, масштаба и образа жизни.', format: 'Индивидуальный', floors: 'По задаче', plan: 'Авторская' }
];

const tabs = document.querySelectorAll('.project-tabs button');
const fields = {
  num: document.querySelector('#project-num'), type: document.querySelector('#project-type'),
  title: document.querySelector('#project-title'), description: document.querySelector('#project-description'),
  format: document.querySelector('#project-format'), floors: document.querySelector('#project-floors'), plan: document.querySelector('#project-plan')
};

tabs.forEach((tab) => tab.addEventListener('click', () => {
  const index = Number(tab.dataset.project); const project = projects[index];
  tabs.forEach((item) => item.classList.remove('active')); tab.classList.add('active');
  fields.num.textContent = String(index + 1).padStart(2, '0'); fields.type.textContent = project.type;
  fields.title.innerHTML = project.title; fields.description.textContent = project.description;
  fields.format.textContent = project.format; fields.floors.textContent = project.floors; fields.plan.textContent = project.plan;
}));

const menuButton = document.querySelector('.menu-button');
const nav = document.querySelector('.nav');
menuButton.addEventListener('click', () => {
  const open = nav.classList.toggle('open'); menuButton.setAttribute('aria-expanded', open);
});
nav.querySelectorAll('a').forEach((link) => link.addEventListener('click', () => { nav.classList.remove('open'); menuButton.setAttribute('aria-expanded', 'false'); }));

const dialog = document.querySelector('.contact-dialog');
document.querySelectorAll('.js-open-contact').forEach((button) => button.addEventListener('click', () => dialog.showModal()));
document.querySelector('.dialog-close').addEventListener('click', () => dialog.close());
dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
dialog.querySelector('form').addEventListener('submit', (event) => {
  event.preventDefault(); event.currentTarget.hidden = true; dialog.querySelector('.form-success').hidden = false;
});

const observer = new IntersectionObserver((entries) => entries.forEach((entry) => {
  if (entry.isIntersecting) { entry.target.animate([{ opacity: 0, transform: 'translateY(28px)' }, { opacity: 1, transform: 'translateY(0)' }], { duration: 700, easing: 'cubic-bezier(.2,.7,.2,1)', fill: 'both' }); observer.unobserve(entry.target); }
}), { threshold: .12 });
document.querySelectorAll('.manifesto-grid, .metric-row, .section-head, .system-layout, .comfort-cards, .process ol').forEach((element) => observer.observe(element));
