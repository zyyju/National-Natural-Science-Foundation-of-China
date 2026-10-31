"""Generate editable PPTX, SVG diagrams and a shared-layout PDF preview.
Requirements: python-pptx, Pillow, fonttools. The PDF preview uses Pillow.
Run from any directory: python 生成源码.py
"""
from pathlib import Path
import json,math,re,html
from functools import lru_cache
from PIL import Image,ImageDraw,ImageFont
from fontTools.ttLib import TTFont
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.oxml.xmlchemy import OxmlElement

OUT=Path(__file__).resolve().parent
DATA=json.loads((OUT/'汇报内容.json').read_text())
ASSETS=OUT/'配图';ASSETS.mkdir(exist_ok=True)
FONT='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
BOLD='/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'
FALL='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FALLB='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
CMAP=TTFont(FONT,fontNumber=0).getBestCmap()
NAVY='17314A';TEAL='008786';GOLD='D29650';RED='BD5B53';INK='294356';MUTED='718393';PALE='EAF4F3';BG='F4F7FA';LINE='DCE5EC';WHITE='FFFFFF'

@lru_cache(None)
def fonts(size,bold=False):
 return ImageFont.truetype(BOLD if bold else FONT,round(size)),ImageFont.truetype(FALLB if bold else FALL,round(size))
def width(t,pt,bold=False):
 a,b=fonts(pt*96/72,bold)
 return sum((a if ord(ch) in CMAP else b).getlength(ch) for ch in t)/96
def wrap(t,pt,w,bold=False):
 out=[]
 for para in t.split('\n'):
  line=''
  tokens=re.findall(r'[A-Za-z0-9₀₁₂₃₄₅₆₇₈₉ₑₓᵧ⁻¹²³]+(?:[/–−.\-][A-Za-z0-9₀₁₂₃₄₅₆₇₈₉ₑₓᵧ⁻¹²³]+)*|.',para)
  for token in tokens:
   if line and width(line+token,pt,bold)>w:out.append(line.rstrip());line=token.lstrip()
   else:line+=token
  out.append(line)
 return '\n'.join(out)

class Scene:
 def __init__(self,w=13.333,h=7.5,bg=BG):self.w=w;self.h=h;self.bg=bg;self.ops=[]
 def box(self,x,y,w,h,fill=WHITE,stroke=None,r=.1):self.ops.append(dict(k='box',x=x,y=y,w=w,h=h,fill=fill,stroke=stroke,r=r))
 def circle(self,x,y,d,fill=TEAL,stroke=None):self.ops.append(dict(k='circle',x=x,y=y,w=d,h=d,fill=fill,stroke=stroke))
 def line(self,x1,y1,x2,y2,color=TEAL,lw=1.5):self.ops.append(dict(k='line',x=x1,y=y1,w=x2-x1,h=y2-y1,color=color,lw=lw))
 def arrow(self,x1,y1,x2,y2,color=TEAL,lw=1.5):
  self.line(x1,y1,x2,y2,color,lw)
  angle=math.atan2(y2-y1,x2-x1);l=.105 if self.w<20 else 9
  self.line(x2,y2,x2-l*math.cos(angle-.55),y2-l*math.sin(angle-.55),color,lw)
  self.line(x2,y2,x2-l*math.cos(angle+.55),y2-l*math.sin(angle+.55),color,lw)
 def text(self,t,x,y,w,h,pt=18,color=INK,bold=False,wrapped=True):
  # Diagram coordinates are converted to inches before text is created.
  if wrapped:t=wrap(t,pt,w-.02,bold)
  self.ops.append(dict(k='text',text=t,x=x,y=y,w=w,h=h,pt=pt,color=color,bold=bold))
 def embed(self,other,x,y,w,h):
  sx=w/other.w;sy=h/other.h
  for old in other.ops:
   q=dict(old);q['x']=x+old['x']*sx;q['y']=y+old['y']*sy;q['w']=old['w']*sx;q['h']=old['h']*sy
   if q['k']=='text':q['pt']=old['pt']*sx
   if q['k']=='box':q['r']=old['r']*sx
   self.ops.append(q)

def visual(kind,labels):
 d=Scene(6,4.5,WHITE)
 def box(x,y,w,h,t,fill=PALE,col=TEAL,pt=16):
  d.box(x,y,w,h,fill,r=.12);d.text(t,x+.15,y+.16,w-.3,h-.22,pt,col,True)
 def node(x,y,col=TEAL,s=.22):d.circle(x-s/2,y-s/2,s,col)
 def atoms(x,y,w,h,accent=0):
  for row in range(3):
   pts=[]
   for j in range(4):
    xx=x+.25+j*(w-.5)/3;yy=y+.25+row*(h-.5)/2+(0.07 if row==1 and j==1 else 0)
    pts.append((xx,yy))
   for a,b in zip(pts,pts[1:]):d.line(*a,*b,LINE,1.4)
   for j,(xx,yy) in enumerate(pts):node(xx,yy,GOLD if (j+row)%5==accent else TEAL,.16)
  d.line(x+.25,y+.25,x+.25,y+h-.25,LINE,1)
  d.line(x+w-.25,y+.25,x+w-.25,y+h-.25,LINE,1)
 if kind in ['cover','lattice']:
  atoms(.7,.8,4.6,2.4)
  d.arrow(1.7,.35,1.7,.95,GOLD);d.text('Sm位 → Ca',.45,.05,2.5,.4,16,TEAL,True)
  d.arrow(4.4,3.75,4.05,2.9,GOLD);d.text('Bi位 → Ge',3.55,3.8,2.2,.4,16,GOLD,True)
  d.circle(2.48,1.47,.8,WHITE,TEAL);d.text('局域\n畸变',2.62,1.58,.66,.65,12,TEAL,True)
  d.text('位点与键合关系概念示意',.75,4.15,4.8,.3,11,MUTED)
 elif kind in ['chain','tasks','ladder','gaps']:
  n=len(labels);h=.68 if n==4 else .87
  for i,lab in enumerate(labels):
   y=.2+i*(4.05/n)
   d.circle(.35,y+.14,.34,TEAL);d.text(str(i+1),.43,y+.17,.24,.3,12,WHITE,True)
   box(1.05,y,4.25,h,lab,PALE if i%2==0 else BG,TEAL if i%2==0 else NAVY,19)
   if i<n-1:d.arrow(3.2,y+h+.05,3.2,y+4.05/n-.05)
 elif kind=='device':
  d.box(.45,1.1,1.05,2.0,RED,r=.08);d.text('热端',.61,1.85,.74,.5,22,WHITE,True)
  d.box(4.5,1.1,1.05,2.0,NAVY,r=.08);d.text('冷端',4.65,1.85,.75,.5,22,WHITE,True)
  for i in range(5):d.box(2+i*.4,1.4,.23,1.45,TEAL if i%2==0 else GOLD,r=.04)
  d.arrow(1.62,2.05,2.0,2.05,RED,2);d.arrow(4.08,2.05,4.4,2.05,TEAL,2)
  d.line(2.1,2.95,2.1,3.55,NAVY);d.line(2.1,3.55,4,3.55,NAVY);d.line(4,3.55,4,2.95,NAVY)
  d.circle(2.85,3.31,.5,GOLD);d.text('电',2.97,3.4,.3,.3,13,WHITE,True)
  d.text('PF = S²σ',2.03,.35,2.55,.6,26,NAVY,True);d.text('温差驱动热电转换',1.55,4.0,4.3,.4,17,TEAL,True)
 elif kind=='balance':
  box(.25,.45,2.6,1.2,'带边收益',PALE,TEAL,22);box(3.15,.45,2.6,1.2,'散射代价','F9EFE8',RED,22)
  d.text('能谷 / 色散',.57,1.98,2.2,.5,17,TEAL);d.text('迁移率损失',3.47,1.98,2.2,.5,17,RED)
  d.arrow(1.55,2.65,2.4,3.25);d.arrow(4.4,2.65,3.6,3.25,RED)
  box(1.55,3.3,3, .82,'PF变化：待验证',BG,NAVY,20)
 elif kind=='timeline':
  d.line(.8,.4,.8,4.0,TEAL,3)
  for i,lab in enumerate(labels):
   y=.25+i*1.02;d.circle(.65,y+.15,.3,TEAL if i<3 else GOLD);box(1.35,y,4.25,.75,lab,BG,NAVY,18)
 elif kind in ['questions','hypothesis','innovation']:
  if kind=='questions':
   box(.25,.45,2.55,1.05,'结构起因',PALE,TEAL,23);box(3.2,.45,2.55,1.05,'输运条件',BG,NAVY,23)
   d.arrow(1.55,1.85,2.4,2.6);d.arrow(4.4,1.85,3.6,2.6)
   box(1.15,2.7,3.75,.95,'局域畸变 → 散射',PALE,TEAL,21)
   d.text('固定T、p与缺陷条件的成组对照',.55,4.0,5.2,.35,13,MUTED)
  elif kind=='hypothesis':
   box(.2,.3,2.55,.85,'化学取代',PALE,TEAL,20);box(3.25,.3,2.55,.85,'几何弛豫',BG,NAVY,20)
   d.arrow(1.5,1.35,2.35,2);d.arrow(4.5,1.35,3.65,2)
   box(1.3,2.1,3.4,.8,'带边态 / 电子—声子耦合',PALE,TEAL,17)
   d.arrow(3,3.05,3,3.5);box(1.3,3.6,3.4,.65,'散射与PF：独立计算检验',BG,NAVY,16)
  else:
   for i,(left,right) in enumerate([('化学 / 几何','散射作用路径'),('散射 / 补偿','PF成立条件')]):
    y=.6+i*2;box(.18,y,2.35,1.1,left,PALE,TEAL,20);d.arrow(2.7,y+.52,3.25,y+.52);box(3.4,y,2.42,1.1,right,BG,NAVY,18)
 elif kind in ['factorial','thermal']:
  for i,lab in enumerate(labels):
   x=.22+(i%2)*3.0;y=.15+(i//2)*2.15;d.box(x,y,2.78,1.95,BG,r=.1)
   d.text(lab,x+.15,y+.14,2.49,.46,15,NAVY,True);atoms(x+.3,y+.72,2.15,.85,i)
  if kind=='factorial':d.text('ΔQint = QCa+Ge − QCa − QGe + Q₀',.3,4.28,5.6,.25,11,TEAL,True)
 elif kind=='paired':
  labs=['母相 / 固定几何','取代 / 固定几何','母相 / 弛豫几何','取代 / 弛豫几何']
  for i,lab in enumerate(labs):
   x=.15+(i%2)*3.15;y=.4+(i//2)*2.2;box(x,y,2.55,1.05,lab,PALE if i%2 else BG,NAVY,15)
  d.arrow(2.78,.95,3.18,.95);d.arrow(2.78,3.15,3.18,3.15)
  d.arrow(1.4,1.65,1.4,2.4);d.arrow(4.55,1.65,4.55,2.4)
  d.text('横向：化学贡献；纵向：几何贡献',.4,4.03,5.25,.4,14,TEAL,True)
 elif kind=='scattering':
  for i,lab in enumerate(labels[:3]):
   x=.15+i*2.0;box(x,.35,1.8,1.05,lab,PALE if i<2 else 'F9EFE8',TEAL if i<2 else RED,15);d.arrow(x+.9,1.65,3,2.5)
  d.circle(2.25,2.28,1.5,NAVY);d.text('散射率\nτ⁻¹',2.51,2.57,1.08,.85,20,WHITE,True)
  d.arrow(3,3.9,3,4.2);d.text('迁移率 / σ / PF',1.45,4.14,3.7,.35,17,TEAL,True)
 elif kind=='aicycle':
  positions=[(.2,.3),(3.38,.3),(3.38,2.9),(.2,2.9)]
  for i,((x,y),lab) in enumerate(zip(positions,labels)):box(x,y,2.42,1.0,lab,PALE if i%2==0 else BG,TEAL if i%2==0 else NAVY,16)
  for (a,b) in [((2.8,.8),(3.2,.8)),((4.6,1.55),(4.6,2.65)),((3.15,3.4),(2.85,3.4)),((1.4,2.65),(1.4,1.55))]:d.arrow(*a,*b)
  d.circle(2.32,1.67,1.35,NAVY);d.text('AI\n辅助选点',2.54,1.9,1.05,.85,17,WHITE,True)
 elif kind=='validation':
  for i,(label,col) in enumerate(zip(['训练组','验证组','独立测试'],[TEAL,GOLD,NAVY])):
   y=.5+i*1.05;d.text(label,.2,y+.1,1.4,.5,17,col,True)
   for j in range(5):d.box(1.85+j*.7,y,.52,.55,col,r=.06)
  d.text('同一构型家族归同一组',1.15,3.8,4.9,.5,19,TEAL,True)
  d.text('主动选点 ↔ 随机选点：相同计算预算',.4,4.23,5.5,.28,12,MUTED)
 elif kind=='roadmap':
  pos=[(.2,.35),(2.22,.35),(4.24,.35),(4.24,2.8),(2.22,2.8),(.2,2.8)]
  labs=['母相基准','四组配对','AI选点','显式散射','机制分解','独立验证']
  for i,((x,y),lab) in enumerate(zip(pos,labs)):box(x,y,1.58,1.1,lab,PALE if i%2 else BG,NAVY,16)
  for a,b in [((1.88,.9),(2.12,.9)),((3.9,.9),(4.14,.9)),((5.03,1.7),(5.03,2.55)),((4.14,3.35),(3.9,3.35)),((2.12,3.35),(1.88,3.35))]:d.arrow(*a,*b)
  d.text('失败构型与反例保留，全流程可回溯',.5,4.17,5.2,.33,13,TEAL,True)
 elif kind=='papers':
  paperlabels=['NbCoSb','Sb–Br共掺\nBi₂Se₃','Hf/Sb共掺\nZrNiSn']
  for i,(lab,journal) in enumerate(zip(paperlabels,['ACS AEM · 2023','RSC Adv. · 2022','Sci. Rep. · 2017'])):
   x=.15+i*2;d.box(x,.45,1.8,3.35,WHITE,LINE,r=.08);d.box(x,.45,1.8,.12,TEAL,r=0)
   d.text(lab,x+.15,.82,1.5,1.05,16,NAVY,True)
   for j in range(4):d.line(x+.18,2.1+j*.28,x+1.6,2.1+j*.28,LINE,1.3)
   d.text(journal,x+.14,3.35,1.53,.3,10,TEAL,True)
  d.text('成果主题示意；非原始论文封面',.8,4.1,4.7,.35,12,MUTED)
 elif kind in ['feasibility','deliverables','references']:
  for i,lab in enumerate(labels):
   x=.23+(i%2)*3.0;y=.3+(i//2)*2.12;d.box(x,y,2.56,1.62,PALE if i%2 else BG,r=.12)
   # Simple document/network emblem supplies a visual anchor.
   d.box(x+.19,y+.22,.38,.5,WHITE,TEAL,r=.02)
   d.line(x+.27,y+.39,x+.49,y+.39,TEAL,1);d.line(x+.27,y+.54,x+.49,y+.54,TEAL,1)
   d.text(lab,x+.2,y+.99,2.21,.44,18,NAVY,True)
 elif kind in ['risk','review','funding']:
  pairs={
   'risk':[('4f / 磁态','合理模型比较'),('稳定性不足','可达组分'),('散射代价','阈值与反例'),('AI无增益','覆盖采样')],
   'review':[('对象依据','Ca/Ge文献'),('电导率可信','显式散射'),('AI可行性','分组 / 基线'),('工作量','分层与退路')],
   'funding':[('已有项目','省基金任务书'),('已完成工作','成果与输入输出'),('本申请新增','机制问题与任务'),('经费 / 工作量','逐项核对边界')]
  }[kind]
  for i,(a,b) in enumerate(pairs):
   y=.25+i*1.05;box(.15,y,2.25,.7,a,BG,NAVY,16);d.arrow(2.62,y+.35,3.12,y+.35);box(3.3,y,2.55,.7,b,PALE,TEAL,15)
 elif kind=='gantt':
  for i,yr in enumerate(['2028','2029','2030']):d.text(yr,2.05+i*1.2,.25,1.1,.5,19,TEAL,True)
  for j,(lab,k) in enumerate([('基准与数据',0),('散射机制',1),('验证与归档',2)]):
   y=1.23+j*1.05;d.text(lab,.18,y+.12,1.7,.4,15,NAVY,True)
   for i in range(3):d.box(2.0+i*1.2,y,1.05,.6,TEAL if i==k else BG,r=.05)
  d.text('时间安排为拟议计划，非实际进度',.55,4.15,5.3,.3,12,MUTED)
 elif kind=='closing':
  for i,lab in enumerate(['局域畸变','载流子散射','PF条件']):box(.32+i*1.95,1.0,1.72,1.0,lab,PALE,NAVY,17)
  d.arrow(2.1,1.48,2.24,1.48);d.arrow(4.05,1.48,4.18,1.48)
  box(1.65,3.05,2.8,.8,'AI辅助获得关键证据',NAVY,WHITE,17)
  d.arrow(3,2.9,3,2.25)
 else:raise ValueError(kind)
 return d

def render(scene,pixels=1600):
 k=pixels/scene.w;im=Image.new('RGB',(pixels,round(scene.h*k)),'#'+scene.bg);dr=ImageDraw.Draw(im)
 for q in scene.ops:
  x=q['x']*k;y=q['y']*k;w=q['w']*k;h=q['h']*k
  if q['k']=='box':dr.rounded_rectangle((x,y,x+w,y+h),radius=max(1,q['r']*k),fill='#'+q['fill'],outline='#'+q['stroke'] if q['stroke'] else None,width=1)
  elif q['k']=='circle':dr.ellipse((x,y,x+w,y+h),fill='#'+q['fill'],outline='#'+q['stroke'] if q['stroke'] else None,width=2)
  elif q['k']=='line':dr.line((x,y,x+w,y+h),fill='#'+q['color'],width=max(1,round(q['lw']*k/72)))
  elif q['k']=='text':
   fs=q['pt']*k/72;f,fb=fonts(fs,q['bold']);lh=fs*1.24
   for i,t in enumerate(q['text'].split('\n')):
    xp=x;baseline=y+i*lh+fs*.91
    for ch in t:
     ff=f if ord(ch) in CMAP else fb;dr.text((round(xp),round(baseline)),ch,font=ff,fill='#'+q['color'],anchor='ls');xp+=ff.getlength(ch)
 return im

def svg(scene):
 k=100;z=[f'<svg xmlns="http://www.w3.org/2000/svg" width="{scene.w*k}" height="{scene.h*k}" viewBox="0 0 {scene.w*k} {scene.h*k}">',f'<rect width="100%" height="100%" fill="#{scene.bg}"/>']
 for q in scene.ops:
  x=q['x']*k;y=q['y']*k;w=q['w']*k;h=q['h']*k
  if q['k']=='box':z.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{q["r"]*k}" fill="#{q["fill"]}" stroke="#{q["stroke"] or q["fill"]}"/>')
  elif q['k']=='circle':z.append(f'<ellipse cx="{x+w/2}" cy="{y+h/2}" rx="{w/2}" ry="{h/2}" fill="#{q["fill"]}" stroke="#{q["stroke"] or q["fill"]}"/>')
  elif q['k']=='line':z.append(f'<line x1="{x}" y1="{y}" x2="{x+w}" y2="{y+h}" stroke="#{q["color"]}" stroke-width="{q["lw"]*k/72}"/>')
  elif q['k']=='text':
   fs=q['pt']*k/72
   for i,t in enumerate(q['text'].split('\n')):z.append(f'<text x="{x}" y="{y+fs*.91+i*fs*1.24}" font-family="Noto Sans CJK SC,DejaVu Sans,sans-serif" font-size="{fs}" font-weight="{700 if q["bold"] else 400}" fill="#{q["color"]}">{html.escape(t)}</text>')
 return '\n'.join(z+['</svg>'])

def to_ppt(scene,slide):
 for q in scene.ops:
  x,y,w,h=[Inches(q[z]) for z in ['x','y','w','h']]
  if q['k'] in ['box','circle']:
   sh=slide.shapes.add_shape(MSO_SHAPE.OVAL if q['k']=='circle' else MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h);sh.fill.solid();sh.fill.fore_color.rgb=RGBColor.from_string(q['fill'])
   if q.get('stroke'):sh.line.color.rgb=RGBColor.from_string(q['stroke']);sh.line.width=Pt(.8)
   else:sh.line.fill.background()
   if q['k']=='box':
    try:sh.adjustments[0]=min(.15,q['r']/min(q['w'],q['h']))
    except Exception:pass
  elif q['k']=='line':
   sh=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,x,y,x+w,y+h);sh.line.color.rgb=RGBColor.from_string(q['color']);sh.line.width=Pt(q['lw'])
  elif q['k']=='text':
   sh=slide.shapes.add_textbox(x,y,w,h);tf=sh.text_frame;tf.clear();tf.word_wrap=False;tf.margin_top=tf.margin_bottom=tf.margin_left=tf.margin_right=0
   for i,t in enumerate(q['text'].split('\n')):
    p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.font.name='Microsoft YaHei';p.font.size=Pt(q['pt']);p.font.bold=q['bold'];p.font.color.rgb=RGBColor.from_string(q['color']);p.space_before=p.space_after=Pt(0);p.line_spacing=1.24
    group='';fallback=None
    for ch in t:
     f=ord(ch) not in CMAP
     if fallback is not None and f!=fallback:
      r=p.add_run();r.text=group
      if fallback:r.font.name='Cambria'
      group=''
     group+=ch;fallback=f
    if group:
     r=p.add_run();r.text=group
     if fallback:r.font.name='Cambria'
    for run in p.runs:
     ea=OxmlElement('a:ea');ea.set('typeface','Microsoft YaHei');run._r.get_or_add_rPr().append(ea)

R=Presentation();R.slide_width=Inches(13.333);R.slide_height=Inches(7.5)
R.core_properties.title=DATA['title'];R.core_properties.author='张菊';R.core_properties.subject='2027青年基金项目论证汇报｜AI辅助商务版'
images=[];scenes=[];audit=[]
for i,content in enumerate(DATA['slides'],1):
 sc=Scene();sl=R.slides.add_slide(R.slide_layouts[6]);sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor.from_string(BG)
 sc.box(0,0,13.333,.12,NAVY,r=0)
 sc.text('NSFC 2027  /  青年科学基金项目（C类）',.72,.34,10,.28,10,TEAL,True)
 sc.text(f'{i:02} / {len(DATA["slides"]):02}',11.7,.34,.95,.29,10,MUTED,True)
 if i==1:
  sc.text('AI辅助的Ca/Ge双位点取代\nSmMg₂Bi₂局域畸变与\n载流子散射机制研究',.75,1.05,7.1,2.2,30,NAVY,True)
  sc.box(.76,3.55,1.04,.055,TEAL,r=0)
  for j,p in enumerate(content['points']):sc.text(p,.79,4.0+j*.65,6.35,.55,18,INK,j==0)
  dg=visual(content['diagram'],content['labels']);sc.embed(dg,7.28,1.95,5.55,4.16)
  sc.text('AI辅助选点 · 第一性原理验证 · 散射机制解释',.79,6.45,11.5,.46,17,TEAL,True)
 else:
  headsize=27
  if len(wrap(content['title'],headsize,11.85,True).split('\n'))>1:headsize=25
  sc.text(content['title'],.75,.86,11.85,.65,headsize,NAVY,True)
  sc.box(.76,1.72,max(1.07,width(content['status'],10,True)+.22),.35,PALE,r=.06);sc.text(content['status'],.86,1.79,4.8,.3,10,TEAL,True)
  for j,p in enumerate(content['points']):
   y=2.35+j*.96;sc.circle(.8,y+.09,.11,TEAL);sc.text(p,1.06,y,5.0,.8,19,INK,j==0)
  sc.box(6.52,2.18,6.08,4.56,WHITE,LINE,r=.13)
  dg=visual(content['diagram'],content['labels']);sc.embed(dg,6.56,2.22,6.0,4.5)
 sc.line(.75,6.99,12.59,6.99,LINE,.8)
 source=content['source'] or '图为研究设计/概念示意，非已完成计算结果。'
 sc.text(source,.77,7.1,11.65,.26,8.2,MUTED)
 sl.notes_slide.notes_text_frame.text=content['note']+'\n\n状态：'+content['status']+'。配图：原创可编辑概念示意；不代表已计算数据。\n来源：'+source
 to_ppt(sc,sl);im=render(sc);images.append(im);scenes.append(sc)
 # Editable diagrams are also provided separately as SVG and PNG.
 (ASSETS/f'{i:02}_{content["diagram"]}.svg').write_text(svg(dg))
 render(dg,960).save(ASSETS/f'{i:02}_{content["diagram"]}.png')
 im.save(OUT/f'页面_{i:02}.png')
 for op in sc.ops:
  if op['k']=='text':
   needed=len(op['text'].split('\n'))*op['pt']/72*1.24
   if needed>op['h']+.09:audit.append(dict(page=i,text=op['text'][:40],need=round(needed,2),box=op['h']))

BASE='2027青年基金_AI辅助局域畸变与载流子散射机制_商务汇报'
R.save(OUT/(BASE+'.pptx'))
images[0].save(OUT/(BASE+'_预览.pdf'),save_all=True,append_images=images[1:],resolution=144)
cols=4;tw=400;th=225;gap=18;rows=math.ceil(len(images)/cols)
sheet=Image.new('RGB',(cols*(tw+gap)+gap,rows*(th+gap+28)+gap),'#E3E9EF');dr=ImageDraw.Draw(sheet)
ff,_=fonts(13)
for i,im in enumerate(images):
 x=gap+(i%cols)*(tw+gap);y=gap+(i//cols)*(th+gap+28);sheet.paste(im.resize((tw,th)),(x,y))
 caption=DATA['slides'][i]['title'][:23].translate(str.maketrans('₀₁₂₃₄₅₆₇₈₉','0123456789'))
 dr.text((x,y+th+4),f'{i+1:02}  '+caption,font=ff,fill='#'+NAVY)
sheet.save(OUT/'全套PPT总览.png')
(OUT/'版式检查.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2))
print('已生成',len(R.slides),'页可编辑PPT；配图SVG/PNG各',len(images),'张；PDF同版式预览。')
print('TEXT_HEIGHT_WARNINGS',json.dumps(audit,ensure_ascii=False))
