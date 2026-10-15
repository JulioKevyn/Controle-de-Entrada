from PIL import Image, ImageDraw, ImageFilter
import numpy as np, io
S=4
orig=Image.open('linkedin/edicao/original2.png').convert('RGB')
im=orig.resize((orig.width*S,orig.height*S),Image.LANCZOS); W,H=im.size
# quad da tela na foto original (168x300) -> escala S
q=[(1.0,109.0),(105.8,107.4),(106.8,165.4),(8.4,177.2)]
dst=[(x*S,y*S) for x,y in q]
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
gy,gx=np.mgrid[0:sh,0:sw]
g=1-0.28*np.clip((((gx-sw*0.45)/(sw*0.62))**2+((gy-sh*0.4)/(sh*0.7))**2),0,1.4)
gw=np.asarray(Image.fromarray((np.clip(g,0,1)*255).astype(np.uint8)).transform((W,H),Image.PERSPECTIVE,c,Image.BICUBIC)).astype(np.float32)/255.0
qm=Image.new('L',(W,H),0); ImageDraw.Draw(qm).polygon(dst,fill=255)
qs=np.asarray(qm.filter(ImageFilter.GaussianBlur(1.6))).astype(np.float32)/255.0
w=np.asarray(warp.filter(ImageFilter.GaussianBlur(2.2))).astype(np.float32)
w=w*gw[...,None]*0.82*np.array([0.94,1.0,1.03])
yy,xx=np.mgrid[0:H,0:W]
w=w*(1-0.012*np.sin(2*np.pi*(yy*0.93+xx*0.12)/9.5))[...,None]
rng=np.random.default_rng(5); w=w+rng.normal(0,3.0,w.shape)
refl=np.clip(1-(xx*0.5+yy*0.9-420)/260,0,1)*np.clip((xx*0.5+yy*0.9-120)/160,0,1)
w=w+refl[...,None]*16
base=np.asarray(im).astype(np.float32)
out=base*(1-qs[...,None])+w*qs[...,None]
bl=np.asarray(Image.fromarray(np.clip(w*qs[...,None],0,255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(14))).astype(np.float32)
out=out+bl*0.12*(1-qs[...,None])
big=Image.fromarray(np.clip(out,0,255).astype(np.uint8))
nat=big.resize(orig.size,Image.LANCZOS)
buf=io.BytesIO(); nat.save(buf,'JPEG',quality=88); buf.seek(0); nat=Image.open(buf).convert('RGB')
nat.save('linkedin/edicao/caixa2-future-pdv-168x300.jpg',quality=95,subsampling=0)
big.save('linkedin/edicao/caixa2-future-pdv-4x.jpg',quality=92,subsampling=0)
big.crop((0,330,W,780)).save('/tmp/check2.png'); print(nat.size,big.size)
