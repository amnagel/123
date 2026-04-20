from fastapi import FastAPI,UploadFile as UF,File as F,Form as FM
from fastapi.responses import StreamingResponse as SR
import pandas as pd,io,pypandoc,json,os
from sympy import symbols as s,latex as l,Matrix as m

app=FastAPI()

def tc(d,p):
    if 'X' not in d.columns:
        for i in range(len(d)):
            if any(str(v).strip().upper() == 'X' for v in d.iloc[i].values):
                d.columns = d.iloc[i]
                d = d.iloc[i+1:].reset_index(drop=True)
                break

    d.columns = [str(c).strip().upper() for c in d.columns]

    if not {'X', 'Y', 'Z'}.issubset(d.columns):
        raise ValueError(f"Колонки X,Y,Z не найдены. В файле есть: {list(d.columns)}")

    c = 1 + (p.get('M', 0) * 1e-6)
    d['NX'] = pd.to_numeric(d['X'], errors='coerce') * c + p.get('DX', 0)
    d['NY'] = pd.to_numeric(d['Y'], errors='coerce') * c + p.get('DY', 0)
    d['NZ'] = pd.to_numeric(d['Z'], errors='coerce') * c + p.get('DZ', 0)
    return d

def gm(d,t,p):
    q1,q2,q3,w1,w2,w3,e1,e2,e3,r=s('X Y Z \\Delta\\ X \\Delta\\ Y \\Delta\\ Z \\omega_X \\omega_Y \\omega_Z m')
    u=m(s('X Y Z'))
    i=m([q1,q2,q3])
    o=m([w1,w2,w3])
    a=m([[1,e3,-e2],[-e3,1,e1],[e2,-e1,1]])
    f=f"{l(u)}=(1+{l(r)}) {l(a)} {l(i)} + {l(o)}"

    res=f"# Отчет\n\n## Формула\n\n$$ {f} $$\n\n"
    res+="| Точка | Исх X | Исх Y | Исх Z | Новая X | Новая Y | Новая Z |\n"
    res+="|---|---|---|---|---|---|---|\n"

    for i in range(len(d)):
        r1,r2=d.iloc[i],t.iloc[i]
        n=r1.get('NAME', r1.get('Name', f'P{i}'))
        res+=f"| {n} | {r1['X']} | {r1['Y']} | {r1['Z']} | {r2['NX']:.3f} | {r2['NY']:.3f} | {r2['NZ']:.3f} |\n"
    return res

@app.post("/p")
async def p(file:UF=F(...),params:str=FM(...)):
    try:
        k_raw=json.loads(params)
        k={key.upper(): float(val) for key, val in k_raw.items()}

        b=await file.read()
        if file.filename.endswith('.csv'):
            df=pd.read_csv(io.BytesIO(b))
        else:
            df=pd.read_excel(io.BytesIO(b))

        td=tc(df.copy(),k)
        md=gm(td,td,k)

        pypandoc.convert_text(md, 'docx', format='md', outputfile='res.docx')

        with open('res.docx', 'rb') as f:
            docx_data = f.read()

        os.remove('res.docx')

        out=io.BytesIO(docx_data)
        out.seek(0)
        return SR(out,media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        return SR(io.BytesIO(f"Ошибка сервера: {str(e)}".encode()), status_code=500)
