"""menu.xlsx + photos/ -> static site/. Runs locally for developers and in GitHub Actions."""
from __future__ import annotations
import argparse, hashlib, html, json, math, os, re, shutil, sys, tempfile
from pathlib import Path
from decimal import Decimal, InvalidOperation
from string import Template
from urllib.parse import urlparse, quote
from openpyxl import load_workbook
from PIL import Image, ImageOps
ROOT=Path(__file__).resolve().parent
HEADERS=['Раздел','Название блюда','Порция / вариант','Цена, ₽','Описание / состав','Фото','Показывать','Код']
CATEGORY_IDS=dict(zip(['Завтраки','Холодные закуски','Горячие закуски','Салаты','Первые блюда','Вторые блюда','Шашлыки','Рыба','Гарниры','Хлеб','Соусы','Детское меню','Напитки','Десерты'],['breakfast','cold','hot','salads','soups','mains','grill','fish','sides','bread','sauces','kids','drinks','desserts']))
class MenuError(ValueError):pass

def text(v):return '' if v is None else str(v).strip()
def esc(v):return html.escape(text(v),quote=True)
def yes_no(v,where,default=True):
 s=text(v).casefold()
 if not s:return default
 if s in ('да','yes','1'):return True
 if s in ('нет','no','0'):return False
 raise MenuError(f'{where}: введите «Да» или «Нет».')
def price(v,row):
 if isinstance(v,bool):raise MenuError(f'Меню, строка {row}: цена должна быть числом.')
 try:n=Decimal(text(v).replace(' ','').replace('\u00a0','').replace(',','.'))
 except InvalidOperation:raise MenuError(f'Меню, строка {row}: укажите цену числом, без ₽.')
 if not n.is_finite() or n<0 or n!=n.quantize(Decimal('.01')):raise MenuError(f'Меню, строка {row}: цена должна быть от 0, не более двух знаков после запятой.')
 return n

def safe_file(folder:Path,name:str,where:str,extensions=None):
 if not name or Path(name).name!=name or any(c in name for c in '/\\:') or name.startswith('.') or len(name)>200:raise MenuError(f'{where}: укажите только имя файла, без папок и ссылок.')
 file=folder/name
 try:
  if not file.is_file() or file.resolve().parent!=folder.resolve():raise OSError()
 except OSError:raise MenuError(f'{where}: не найден файл {name}. Загрузите его в папку {folder.name}.')
 if extensions and file.suffix.casefold() not in extensions:raise MenuError(f'{where}: поддерживаются JPG, PNG и WebP.')
 return file

def read_menu(workbook:Path,root:Path):
 try:wb=load_workbook(workbook,data_only=False)
 except Exception as e:raise MenuError('Не удалось открыть menu.xlsx. Сохраните обычный файл Excel .xlsx и попробуйте снова.') from e
 try:
  for name in ['Меню','Разделы','Настройки','Ссылки']:
   if name not in wb.sheetnames:raise MenuError(f'В Excel отсутствует лист «{name}». Не переименовывайте листы.')
  for sh in wb:
   if sh.max_row>10000 or sh.max_column>100:raise MenuError(f'Лист «{sh.title}» слишком большой: удалите лишнее форматирование за пределами таблицы.')
   for row in sh:
    for cell in row:
     if cell.data_type=='f':raise MenuError(f'Лист «{sh.title}», ячейка {cell.coordinate}: формулы не поддерживаются; вставьте готовое значение.')
  ws=wb['Меню']
  if [text(c.value) for c in ws[1]][:8]!=HEADERS:raise MenuError('Не меняйте заголовки колонок листа «Меню».')
  categories=[];seen=set()
  for row in wb['Разделы'].iter_rows(min_row=2,values_only=True):
   name=text(row[0])
   if not name:continue
   if name.casefold() in seen:raise MenuError(f'Повтор раздела «{name}» на листе «Разделы».')
   seen.add(name.casefold());categories.append({'name':name,'id':CATEGORY_IDS.get(name,'c-'+hashlib.sha256(name.encode()).hexdigest()[:12])})
  allowed={c['name'] for c in categories};dishes=[];ids=set();current=None
  for num,row in enumerate(ws.iter_rows(min_row=2,max_col=8,values_only=True),2):
   if not any(text(v) for v in row):continue
   cat,name,portion,cost,description,photo,visible,code=row
   if not text(name):
    if current is None:raise MenuError(f'Меню, строка {num}: у порции нет блюда. Укажите название в первой строке блюда.')
    if any(text(v) for v in [cat,description,photo,visible,code]):raise MenuError(f'Меню, строка {num}: дополнительная порция содержит лишние поля. Заполните только колонки C и D или укажите название нового блюда.')
    current['portions'].append({'label':text(portion),'price':price(cost,num)});continue
   if text(cat) not in allowed:raise MenuError(f'Меню, строка {num}: раздел «{text(cat)}» не найден. Выберите значение из списка разделов.')
   ident=text(code) or 'dish-'+hashlib.sha256((text(cat)+'|'+text(name)).encode()).hexdigest()[:16]
   if not re.fullmatch(r'[a-zA-Z0-9_-]{1,80}',ident):raise MenuError(f'Меню, строка {num}: код содержит недопустимые символы.')
   if ident in ids:raise MenuError(f'Меню, строка {num}: повторяется код «{ident}». У нового скопированного блюда очистите код в колонке H.')
   ids.add(ident)
   if text(photo):safe_file(root/'photos',text(photo),f'Меню, строка {num}',{'.jpg','.jpeg','.png','.webp'})
   current={'id':ident,'category':text(cat),'name':text(name),'description':text(description),'photo':text(photo),'visible':yes_no(visible,f'Меню, строка {num}'),'portions':[{'label':text(portion),'price':price(cost,num)}]};dishes.append(current)
  if not any(d['visible'] for d in dishes):raise MenuError('В меню нет видимых блюд. Добавьте хотя бы одно и выберите «Да» в колонке G.')
  settings={}
  for row in wb['Настройки'].iter_rows(min_row=2,max_col=2,values_only=True):
   if not text(row[0]):continue
   if text(row[0]) in settings:raise MenuError(f'Повтор настройки «{text(row[0])}».')
   settings[text(row[0])]=text(row[1])
  if not settings.get('Название'):raise MenuError('На листе «Настройки» укажите название кафе.')
  settings['draft']=yes_no(settings.get('Предварительная версия'),'Настройки: Предварительная версия')
  for key in ['Фон','Логотип']:
   if settings.get(key):safe_file(root/'assets',settings[key],f'Настройки: {key}',{'.webp','.jpg','.jpeg','.png'})
  site_url=settings.get('Сайт URL','').rstrip('/')
  if site_url:
   u=urlparse(site_url)
   if u.scheme!='https' or not u.hostname or u.query or u.fragment or u.username:raise MenuError('Настройки: «Сайт URL» должен быть настоящим HTTPS-адресом без параметров и #якорей.')
  settings['site_url']=site_url
  telegram_url=settings.get('Telegram URL','')
  if telegram_url:
   u=urlparse(telegram_url)
   if u.scheme!='https' or not u.hostname or u.username:raise MenuError('Настройки: «Telegram URL» должен быть полным HTTPS-адресом.')
  settings['telegram_url']=telegram_url
  links=[]
  for num,row in enumerate(wb['Ссылки'].iter_rows(min_row=2,max_col=2,values_only=True),2):
   label,url=map(text,row)
   if not label and not url:continue
   u=urlparse(url)
   if not label or u.scheme!='https' or not u.hostname or u.username:raise MenuError(f'Ссылки, строка {num}: укажите подпись и полный адрес https://...')
   links.append({'label':label,'url':url})
  return {'categories':categories,'dishes':dishes,'settings':settings,'links':links}
 finally:wb.close()

def money(value):
 s=f'{value:,.2f}'.replace(',',' ').rstrip('0').rstrip('.')
 return s+' ₽'

def read_framing(root):
 """Optional display crops; source photos and full-detail views stay intact."""
 file=Path(root)/'photo-framing.json'
 if not file.exists():return {}
 try:
  if file.stat().st_size>1000000:raise ValueError('Файл больше 1 МБ.')
  data=json.loads(file.read_text(encoding='utf-8-sig'))
  if not isinstance(data,dict) or data.get('version')!=1 or not isinstance(data.get('photos'),dict):raise ValueError('Нужны version: 1 и объект photos.')
  result={}
  for name,values in data['photos'].items():
   if not isinstance(name,str) or Path(name).name!=name or any(c in name for c in '/\\:') or name.startswith('.') or Path(name).suffix.casefold() not in {'.jpg','.jpeg','.png','.webp'}:raise ValueError('Недопустимое имя фотографии.')
   if not isinstance(values,dict) or set(values)-{'zoom','x','y','source_sha256'}:raise ValueError(f'{name}: неизвестная настройка кадра.')
   digest=values.get('source_sha256')
   if digest is not None and (not isinstance(digest,str) or not re.fullmatch(r'[0-9a-f]{64}',digest)):raise ValueError(f'{name}: неверная контрольная сумма исходника.')
   crop={'zoom':values.get('zoom',1),'x':values.get('x',0),'y':values.get('y',0)}
   if any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in crop.values()):raise ValueError(f'{name}: настройки кадра должны быть конечными числами.')
   if not 1<=crop['zoom']<=3:raise ValueError(f'{name}: увеличение должно быть от 1 до 3.')
   limit=(crop['zoom']-1)*50
   if abs(crop['x'])>limit+0.00001 or abs(crop['y'])>limit+0.00001:raise ValueError(f'{name}: смещение оставляет пустой край фотографии.')
   source=Path(root)/'photos'/name
   if digest is not None and source.is_file() and hashlib.sha256(source.read_bytes()).hexdigest()!=digest:
    print(f'ФОТО: {name} изменено. Использован обычный кадр 4:3; приближение для нового фото нужно подобрать отдельно.',file=sys.stderr)
    continue
   result[name]=crop
  return result
 except (OSError,UnicodeError,ValueError,TypeError) as e:raise MenuError(f'photo-framing.json: {e}') from e

def css_number(value):return f'{value:.4f}'.rstrip('0').rstrip('.') or '0'

def build_images(model,root,stage):
 images={};cache={};out=stage/'photos';out.mkdir()
 for dish in model['dishes']:
  if not dish['visible'] or not dish['photo']:continue
  file=safe_file(root/'photos',dish['photo'],'Фото блюда '+dish['name'])
  if file.name not in cache:
   if file.stat().st_size>50*1024*1024:raise MenuError(f'Фото «{file.name}» больше 50 МБ: сначала подготовьте уменьшенную копию.')
   try:
    with Image.open(file) as raw:
     if raw.width*raw.height>80000000:raise MenuError(f'Фото «{file.name}» больше 80 мегапикселей.')
     im=ImageOps.exif_transpose(raw);im=im.convert('RGBA' if 'A' in im.getbands() else 'RGB')
     digest=hashlib.sha256(file.read_bytes()).hexdigest()[:18];names={}
     for size in [720,1200]:
      thumb=im.copy();thumb.thumbnail((size,size),Image.Resampling.LANCZOS);name=digest+f'-{size}.webp';thumb.save(out/name,'WEBP',quality=82,method=4);names[size]=name
     cache[file.name]=names
   except MenuError:raise
   except Exception as e:raise MenuError(f'Не удалось обработать фото «{file.name}». Проверьте, что это JPG, PNG или WebP.') from e
  images[dish['id']]=cache[file.name]
 return images

def render(model,images,root,asset_root=None,framing=None):
 s=model['settings'];cats=[];sections=[];category_list=[];framing=framing or {}
 for c in model['categories']:
  dishes=[d for d in model['dishes'] if d['category']==c['name'] and d['visible']]
  if not dishes:continue
  cats.append(f'<a href="#category-{c["id"]}">{esc(c["name"])}</a>')
  category_list.append(f'<a href="#category-{c["id"]}" data-category-target="category-{c["id"]}"><span>{esc(c["name"])}</span><span class="category-count">{len(dishes)}</span></a>')
  cards=[]
  for index,d in enumerate(dishes):
   if d['id'] in images:
    src=images[d['id']];priority='fetchpriority="high"' if not sections and index<2 else 'loading="lazy"'
    crop=framing.get(d['photo']);zoom=crop['zoom'] if crop else 1
    frame_style=f' style="--photo-zoom:{css_number(zoom)};--photo-x:{css_number(crop["x"])}%;--photo-y:{css_number(crop["y"])}%"' if crop else ''
    sizes=f'(max-width:640px) calc({css_number(100*zoom)}vw - {css_number(28*zoom)}px), (max-width:980px) {css_number(48*zoom)}vw, {css_number(360*zoom)}px'
    visual=f'<div class="photo-frame"{frame_style}><img src="./photos/{src[720]}" srcset="./photos/{src[720]} 720w, ./photos/{src[1200]} 1200w" sizes="{sizes}" alt="{esc(d["name"])}" {priority} decoding="async" width="720" height="600"></div>'
   else:visual=f'<div class="placeholder" aria-hidden="true"><span class="placeholder-label">{esc(c["name"])}</span><span class="placeholder-name">{esc(d["name"])}</span><span class="placeholder-caption">Фото появится позже</span></div>'
   portions=''.join(f'<div><span>{esc(p["label"])}</span><strong>{money(p["price"])}</strong></div>' for p in d['portions'])
   long=len(d['portions'])>4
   search=' '.join([d['name'],d['description'],c['name'],*[p['label'] for p in d['portions']]]).lower().replace('ё','е')
   description=f'<p class="description">{esc(d["description"])}</p>' if d['description'] else ''
   variants_note=f'<p class="variants-note">Виды, объёмы и цены внутри · {len(d["portions"])} вариантов</p>' if long else ''
   cards.append(f'<article id="dish-{d["id"]}" data-search="{esc(search)}"><a class="dish-card" href="#dish-{d["id"]}" data-open="{d["id"]}" aria-haspopup="dialog">{visual}<div class="card-body"><h3>{esc(d["name"])}</h3>{description}<div class="portions {"is-long" if long else ""}">{portions}</div>{variants_note}<span class="more">Подробнее</span></div></a></article>')
  bonus=''
  if c['id']=='mains' and s.get('Бонус к блюду'):bonus=f'<aside class="bonus"><strong>Бонус к блюду</strong><span>{esc(s["Бонус к блюду"])}</span><small>{esc(s.get("Условия бонуса",""))}</small></aside>'
  sections.append(f'<section class="menu-section" id="category-{c["id"]}" data-menu-section><div class="section-head"><h2>{esc(c["name"])}</h2><span>{len(dishes)} поз.</span></div>{bonus}<div class="grid">{"".join(cards)}</div></section>')
 links=''.join(f'<a href="{esc(l["url"])}" target="_blank" rel="noopener noreferrer">{esc(l["label"])}</a>' for l in model['links'])
 phone=s.get('Телефон','');phone_html=f'<a class="contact-phone" href="tel:{esc(re.sub(r"[^+0-9]","",phone))}">{esc(phone)}</a>' if phone else ''
 logo_hash=hashlib.sha256((root/'assets'/s['Логотип']).read_bytes()).hexdigest()[:12] if s.get('Логотип') else ''
 logo=f'<img class="brand-logo" src="./assets/{esc(s["Логотип"])}?v={logo_hash}" alt="Кафе {esc(s['Название'])}" width="180" height="96">' if s.get('Логотип') else ''
 phone_icon='<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M7.2 3.5 4.5 5.2c-1.2.8-.5 5.1 3.4 9.1s8.2 4.7 9.1 3.5l1.7-2.7-4.3-2.2-1.5 1.5a15 15 0 0 1-4.3-4.3l1.5-1.5Z"/></svg>'
 telegram_icon='<svg aria-hidden="true" viewBox="0 0 24 24"><path d="m3 10 18-7-4 18-6-6-4 3 1-6Z"/><path d="m8 12 9-5-6 8"/></svg>'
 header_phone=f'<a class="contact-action phone-action" href="tel:{esc(re.sub(r"[^+0-9]","",phone))}" aria-label="Позвонить в кафе: {esc(phone)}">{phone_icon}<span>{esc(phone)}</span></a>' if phone else ''
 telegram_url=s['telegram_url'] or next((l['url'] for l in model['links'] if l['label'].casefold() in ('telegram','телеграм')), '')
 header_telegram=f'<a class="contact-action telegram-action" href="{esc(telegram_url)}" target="_blank" rel="noopener noreferrer" aria-label="Открыть Telegram кафе">{telegram_icon}<span>Telegram</span></a>' if telegram_url else f'<button type="button" class="contact-action telegram-action is-pending" disabled aria-label="Telegram: ссылка появится позже">{telegram_icon}<span>Telegram <small>скоро</small></span></button>'
 header_contacts=header_phone+header_telegram
 about=''.join(f'<p>{esc(p)}</p>' for p in s.get('О кафе','').split('\n\n') if p.strip())
 delivery=esc(s.get('Доставка','')).replace('\n','<br>')
 meta='<meta name="robots" content="noindex,nofollow">' if s['draft'] else ''
 if s['site_url']:meta+=f'<link rel="canonical" href="{esc(s["site_url"])}/"><meta property="og:url" content="{esc(s["site_url"])}/">'
 asset_root=asset_root or root/'assets'
 css_hash=hashlib.sha256((asset_root/'site.css').read_bytes()).hexdigest()[:12];js_hash=hashlib.sha256((asset_root/'site.js').read_bytes()).hexdigest()[:12]
 return Template((root/'templates/index.html').read_text()).substitute(name=esc(s['Название']),subtitle=esc(s.get('Подпись','')),description=esc(s.get('Описание','')),meta=meta,logo=logo,brand_class="has-logo" if logo else "",header_contacts=header_contacts,tabs=''.join(cats),categories=''.join(category_list),sections=''.join(sections),total=sum(d['visible'] for d in model['dishes']),about=about,phone=phone_html,address=esc(s.get('Адрес','')),hours=esc(s.get('Режим работы','')),delivery=delivery,links=links,css_hash=css_hash,js_hash=js_hash)

def build(root=ROOT,output=None,workbook=None):
 root=Path(root).resolve();dest=Path(output or root/'site').resolve();workbook=Path(workbook or root/'menu.xlsx')
 if dest==root or root.is_relative_to(dest) or dest.is_relative_to(root/'photos') or dest.is_relative_to(root/'assets'):raise MenuError('Недопустимая папка результата.')
 model=read_menu(workbook,root);framing=read_framing(root);dest.parent.mkdir(parents=True,exist_ok=True);stage=Path(tempfile.mkdtemp(prefix='.site-build-',dir=dest.parent));old=None
 try:
  images=build_images(model,root,stage);shutil.copytree(root/'assets',stage/'assets')
  bg=model['settings'].get('Фон','');mobile='background-mobile.webp' if bg=='background.webp' else bg
  background="url('./"+quote(bg,safe='')+"')" if bg else 'none'
  mobile_background="url('./"+quote(mobile,safe='')+"')" if mobile else 'none'
  with (stage/'assets/site.css').open('a') as css:css.write('\nbody{background-image:'+background+'}@media(max-width:640px){body{background-image:'+mobile_background+'}}\n')
  (stage/'index.html').write_text(render(model,images,root,stage/'assets',framing),encoding='utf-8');(stage/'.nojekyll').write_text('')
  settings=model['settings'];url=settings['site_url']
  (stage/'robots.txt').write_text('User-agent: *\n'+('Disallow: /\n' if settings['draft'] else 'Allow: /\n')+(f'Sitemap: {url}/sitemap.xml\n' if url and not settings['draft'] else ''))
  if url:(stage/'sitemap.xml').write_text('<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"><url><loc>'+html.escape(url+'/')+'</loc></url></urlset>')
  if dest.exists():old=dest.with_name(dest.name+'.previous-'+next(tempfile._get_candidate_names()));dest.rename(old)
  try:stage.rename(dest)
  except Exception:
   if old:old.rename(dest)
   raise
  if old:shutil.rmtree(old)
  report={'sections':len({d['category'] for d in model['dishes'] if d['visible']}),'cards':sum(d['visible'] for d in model['dishes']),'portions':sum(len(d['portions']) for d in model['dishes'] if d['visible']),'cardsWithPhotos':len(images),'placeholderCards':[d['name'] for d in model['dishes'] if d['visible'] and d['id'] not in images],'draft':settings['draft']}
  return report
 finally:
  if stage.exists():shutil.rmtree(stage)

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path);parser.add_argument('--workbook',type=Path);args=parser.parse_args()
 try:
  result=build(output=args.output,workbook=args.workbook);print('Готово:',json.dumps(result,ensure_ascii=False));(ROOT/'build-report.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
 except MenuError as e:print('ОШИБКА МЕНЮ:',e,file=sys.stderr);sys.exit(1)
