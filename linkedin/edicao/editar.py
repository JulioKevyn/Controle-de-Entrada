from PIL import Image, ImageDraw, ImageFilter
import numpy as np, io
im=Image.open('linkedin/edicao/original.png').convert('RGB'); W,H=im.size
# 1) remove o "13" clonando o poste (com compensação da inclinação da borda)
a=np.asarray(im).astype(np.float32).copy()
mask=Image.new('L',(W,H),0); ImageDraw.Draw(mask).polygon([(454,176),(521,170),(523,258),(462.5,260)],fill=255)
m=(np.asarray(mask.filter(ImageFilter.GaussianBlur(1.6))).astype(np.float32)/255.0)[...,None]
ys,xs=np.mgrid[0:H,0:W]
src=a[np.clip((ys+80).astype(int),0,H-1),np.clip((xs+8.5).astype(int),0,W-1)]
a=a*(1-m)+src*m
im=Image.fromarray(np.clip(a,0,255).astype(np.uint8))
# 2) tela do Future PDV, com acabamento de foto de monitor
dst=[(123,593),(541,553),(559,813),(139,828)]
scr=Image.open('video-mercado/out/sys/mk-caixa.png').convert('RGB').resize((880,522),Image.LANCZOS)
sw,sh=scr.size
def coeffs(s,d):
    A=[];B=[]
    for (x,y),(u,v) in zip(s,d):
        A.append([x,y,1,0,0,0,-u*x,-u*y]);B.append(u)
        A.append([0,0,0,x,y,1,-v*x,-v*y]);B.append(v)
    return np.linalg.solve(np.array(A,float),np.array(B,float))
c=tuple(coeffs(dst,[(0,0),(sw,0),(sw,sh),(0,sh)]))
warp=scr.transform((W,H),Image.PERSPECTIVE,c,Image.BICUBIC)
# gradiente (centro mais claro, bordas mais escuras) no plano da tela
gy,gx=np.mgrid[0:sh,0:sw]
g=1-0.28*np.clip((((gx-sw*0.45)/(sw*0.62))**2+((gy-sh*0.4)/(sh*0.7))**2),0,1.4)
gw=np.asarray(Image.fromarray((np.clip(g,0,1)*255).astype(np.uint8)).transform((W,H),Image.PERSPECTIVE,c,Image.BICUBIC)).astype(np.float32)/255.0
qm=Image.new('L',(W,H),0); ImageDraw.Draw(qm).polygon(dst,fill=255)
qs=np.asarray(qm.filter(ImageFilter.GaussianBlur(0.9))).astype(np.float32)/255.0
w=np.asarray(warp.filter(ImageFilter.GaussianBlur(0.75))).astype(np.float32)
w=w*gw[...,None]*0.82                       # menos estourada
w=w*np.array([0.94,1.0,1.03])               # leve tom frio/esverdeado de LCD
yy,xx=np.mgrid[0:H,0:W]
w=w*(1-0.028*np.sin(2*np.pi*(yy*0.93+xx*0.12)/2.7))[...,None]   # trama dos pixels/moiré
rng=np.random.default_rng(5); w=w+rng.normal(0,3.2,w.shape)
refl=np.clip(1-(xx*0.5+yy*0.9-700)/300,0,1)*np.clip((xx*0.5+yy*0.9-400)/220,0,1)
w=w+refl[...,None]*20
base=np.asarray(im).astype(np.float32)
dark=np.array([12,16,20],np.float32)
out=base*(1-qs[...,None])+(w*qs[...,None]+dark*0*qs[...,None])
# brilho (bloom) da tela vazando para a moldura
bl=np.asarray(Image.fromarray(np.clip(w*qs[...,None],0,255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(10))).astype(np.float32)
out=out+bl*0.14*(1-qs[...,None])
img=Image.fromarray(np.clip(out,0,255).astype(np.uint8))
# 3) recompressão (mesma "cara" de foto de celular)
buf=io.BytesIO(); img.save(buf,'JPEG',quality=88); buf.seek(0); img=Image.open(buf).convert('RGB')
img.save('linkedin/edicao/caixa-future-pdv.jpg',quality=92,subsampling=0)
img.crop((60,500,640,880)).resize((1160,760)).save('/tmp/check_screen.png')
print(img.size)
