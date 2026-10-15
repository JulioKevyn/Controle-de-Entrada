# Super-resolução com Swin2SR (x2) via ONNX, aplicada duas vezes (x4) com divisão em blocos
import numpy as np, onnxruntime as ort, sys
from PIL import Image
sess=ort.InferenceSession('/tmp/sr/swin2SR-classical-sr-x2-64.onnx')
def sr2(img, tile=96, pad=20):
    a=np.asarray(img.convert('RGB')).astype(np.float32)/255.0
    h,w,_=a.shape; out=np.zeros((h*2,w*2,3),np.float32)
    for y in range(0,h,tile):
        for x in range(0,w,tile):
            y0,x0=max(0,y-pad),max(0,x-pad); y1,x1=min(h,y+tile+pad),min(w,x+tile+pad)
            t=a[y0:y1,x0:x1]; th,tw=t.shape[:2]
            ph,pw=(-th)%8,(-tw)%8
            t=np.pad(t,((0,ph),(0,pw),(0,0)),mode='reflect')
            r=sess.run(None,{'pixel_values':t.transpose(2,0,1)[None]})[0][0].transpose(1,2,0)
            r=r[:th*2,:tw*2]
            oy,ox=(y-y0)*2,(x-x0)*2
            hh,ww=min(tile,h-y)*2,min(tile,w-x)*2
            out[y*2:y*2+hh,x*2:x*2+ww]=r[oy:oy+hh,ox:ox+ww]
    return Image.fromarray((np.clip(out,0,1)*255).astype(np.uint8))
if __name__=='__main__':
    im=Image.open(sys.argv[1]).convert('RGB')
    x2=sr2(im); x4=sr2(x2,tile=128)
    lan=im.resize(x4.size,Image.LANCZOS)
    x4=Image.blend(lan,x4,0.6)
    x4.save(sys.argv[2]); print(im.size,'->',x4.size)
