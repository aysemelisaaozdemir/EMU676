#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun  9 18:56:26 2026

@author: melisaozdemir
"""

#!/usr/bin/env python3
"""
ARP-SA v2 — ALNS Çözücü (Düzeltilmiş Sürüm)
Assembly Routing Problem with Subcontractor Allocation

Düzeltmeler (ara rapor tutarsızlıkları):
  [D1] destroy_worst / destroy_bottleneck → tam q adet (p,tau) çifti kaldırır
  [D2] build_initial → ε-tabanlı aparat uyumu tie-breaking eklendi (rapor s.9 md.3)
  [D3] Rota yeniden inşası → _try_insert içindeki davranış belgelendi
"""
import random, math, time, json, os
from typing import List, Tuple

# ═══════════════════════════════════════════════════════════════
#  PAYLAŞILAN ALTYAPI (tüm instance'lar)
# ═══════════════════════════════════════════════════════════════
TAU    = ['M','E','W']
OPT    = {'M':['WM','sM1','sM2'],'E':['WE','sE1','sE2'],'W':['WW','sW1','sW2']}
INT    = {'M':'WM','E':'WE','W':'WW'}
FIRMS  = ['sM1','sM2','sE1','sE2','sW1','sW2']
FTYPE  = {s:t for t in TAU for s in OPT[t]}
NVEH   = 2

NODES = ['depot','sM1','sM2','sE1','sE2','sW1','sW2']
NI    = {n:i for i,n in enumerate(NODES)}
_TT   = [
    [0.0,1.3,2.1,0.9,1.7,1.5,2.4],
    [1.3,0.0,1.4,1.8,2.3,2.1,2.9],
    [2.1,1.4,0.0,2.5,1.9,2.7,1.6],
    [0.9,1.8,2.5,0.0,1.1,1.3,2.0],
    [1.7,2.3,1.9,1.1,0.0,1.8,1.4],
    [1.5,2.1,2.7,1.3,1.8,0.0,1.2],
    [2.4,2.9,1.6,2.0,1.4,1.2,0.0],
]
def tt(i,j): return _TT[NI[i]][NI[j]]

# ── Instance verisi (patching için modül seviyesi değişkenler) ─
ORDERS = ['P01','P02','P03','P04','P05','P06','P07','P08','P09','P10']
DUE    = {'P01':12,'P02':24,'P03':32,'P04':8,'P05':48,
           'P06':25,'P07':20,'P08':28,'P09':22,'P10':34}
PROC   = {
    'P01':{'M':{'WM':6.5,'sM1':5.2,'sM2':7.8},'E':{'WE':3.8,'sE1':3.1,'sE2':5.4},'W':{'WW':3.2,'sW1':3.9,'sW2':2.7}},
    'P02':{'M':{'WM':8.1,'sM1':6.4,'sM2':9.3},'E':{'WE':5.6,'sE1':4.7,'sE2':6.2},'W':{'WW':4.1,'sW1':3.6,'sW2':4.8}},
    'P03':{'M':{'WM':5.9,'sM1':4.8,'sM2':6.5},'E':{'WE':7.2,'sE1':6.1,'sE2':8.0},'W':{'WW':5.5,'sW1':5.9,'sW2':4.8}},
    'P04':{'M':{'WM':4.3,'sM1':3.7,'sM2':5.1},'E':{'WE':3.2,'sE1':2.8,'sE2':3.9},'W':{'WW':2.8,'sW1':2.4,'sW2':3.3}},
    'P05':{'M':{'WM':9.7,'sM1':8.2,'sM2':11.1},'E':{'WE':6.4,'sE1':5.5,'sE2':7.3},'W':{'WW':6.8,'sW1':7.2,'sW2':6.1}},
    'P06':{'M':{'WM':7.2,'sM1':5.9,'sM2':8.4},'E':{'WE':4.9,'sE1':4.2,'sE2':5.7},'W':{'WW':4.4,'sW1':4.8,'sW2':3.9}},
    'P07':{'M':{'WM':5.1,'sM1':4.3,'sM2':5.8},'E':{'WE':4.1,'sE1':3.5,'sE2':4.7},'W':{'WW':3.6,'sW1':3.2,'sW2':4.1}},
    'P08':{'M':{'WM':7.8,'sM1':6.5,'sM2':8.9},'E':{'WE':5.3,'sE1':4.6,'sE2':6.1},'W':{'WW':5.1,'sW1':5.5,'sW2':4.6}},
    'P09':{'M':{'WM':6.2,'sM1':5.1,'sM2':7.0},'E':{'WE':4.7,'sE1':3.9,'sE2':5.4},'W':{'WW':3.8,'sW1':4.2,'sW2':3.4}},
    'P10':{'M':{'WM':8.5,'sM1':7.1,'sM2':9.6},'E':{'WE':6.8,'sE1':5.7,'sE2':7.5},'W':{'WW':5.9,'sW1':6.3,'sW2':5.2}},
}
FIX  = {
    'P01':{'M':'F2','E':'F1','W':'F3'},'P02':{'M':'F1','E':'F3','W':'F2'},
    'P03':{'M':'F3','E':'F2','W':'F1'},'P04':{'M':'F2','E':'F4','W':'F3'},
    'P05':{'M':'F4','E':'F1','W':'F2'},'P06':{'M':'F1','E':'F2','W':'F4'},
    'P07':{'M':'F3','E':'F4','W':'F1'},'P08':{'M':'F2','E':'F3','W':'F2'},
    'P09':{'M':'F4','E':'F1','W':'F3'},'P10':{'M':'F1','E':'F2','W':'F4'},
}
SU   = {'WM':1.5,'sM1':1.2,'sM2':1.8,'WE':0.8,'sE1':0.6,'sE2':1.0,
        'WW':1.0,'sW1':0.9,'sW2':1.1}

# ═══════════════════════════════════════════════════════════════
#  ÇÖZÜM YAPISI  asgn={p:{tau:s}}, seqs={s:[p,...]}, routes=[[s,...]]
# ═══════════════════════════════════════════════════════════════
def make_empty_seqs():
    return {s:[] for t in TAU for s in OPT[t]}

def copy_sol(a,s,r):
    return ({p:dict(v) for p,v in a.items()},
            {k:list(v) for k,v in s.items()},
            [list(x) for x in r])

# ═══════════════════════════════════════════════════════════════
#  DEĞERLENDİRME
# ═══════════════════════════════════════════════════════════════
def evaluate(asgn, seqs, routes):
    et = {}
    for tau in TAU:
        for s in OPT[tau]:
            seq = seqs.get(s,[])
            t = 0.0
            for k,p in enumerate(seq):
                if k>0 and FIX[seq[k-1]][tau]!=FIX[p][tau]: t+=SU[s]
                t+=PROC[p][tau][s]; et[(p,tau,s)]=t
    ff = {s: max((et.get((p,FTYPE[s],s),0.) for p in seqs.get(s,[])),default=0.)
          for s in FIRMS}
    dep = {}
    for route in routes:
        t,prev=0.,'depot'
        for s in route:
            arr=t+tt(prev,s); dep[s]=max(arr,ff[s]); t,prev=dep[s],s
    r_arr={p:{} for p in ORDERS}
    for p in ORDERS:
        for tau in TAU:
            s=asgn.get(p,{}).get(tau)
            if s is None: r_arr[p][tau]=0.
            elif s==INT[tau]: r_arr[p][tau]=et.get((p,tau,s),0.)
            else: r_arr[p][tau]=dep.get(s,0.)+tt(s,'depot')
    C={p:max(r_arr[p].values()) for p in ORDERS}
    T={p:max(C[p]-DUE[p],0.) for p in ORDERS}
    return sum(T.values()),r_arr,C,T

def obj(a,s,r): return evaluate(a,s,r)[0]

# ═══════════════════════════════════════════════════════════════
#  ROTA İNŞASI
# ═══════════════════════════════════════════════════════════════
def build_routes(needed):
    if not needed: return []
    needed=list(needed); needed.sort(key=lambda s:tt('depot',s))
    grps=[[] for _ in range(NVEH)]
    for i,s in enumerate(needed): grps[i%NVEH].append(s)
    result=[]
    for grp in grps:
        if not grp: continue
        rem,curr,route=list(grp),'depot',[]
        while rem:
            nx=min(rem,key=lambda s:tt(curr,s)); route.append(nx); rem.remove(nx); curr=nx
        result.append(route)
    return result

def needed_firms(asgn):
    return {asgn[p][tau] for p in ORDERS for tau in TAU
            if asgn.get(p,{}).get(tau) and asgn[p][tau]!=INT[tau]}

def needed_firms_partial(asgn):
    f=set()
    for p in ORDERS:
        for tau in TAU:
            s=asgn.get(p,{}).get(tau)
            if s and s!=INT[tau]: f.add(s)
    return f

# ═══════════════════════════════════════════════════════════════
#  BAŞLANGIÇ ÇÖZÜMÜ
#  [D2] ε-tabanlı aparat uyumu tie-breaking (rapor s.9 md.3)
# ═══════════════════════════════════════════════════════════════
def _seq_end(seq,tau,s):
    t=0.
    for k,p in enumerate(seq):
        if k>0 and FIX[seq[k-1]][tau]!=FIX[p][tau]: t+=SU[s]
        t+=PROC[p][tau][s]
    return t

def build_initial():
    """
    EDD sıralı greedy.
    Seçenek değerlendirmesinde iki kriter:
      1) Birincil: tahmini depoya varış süresi (en küçük)
      2) İkincil (ε=0.5h içinde): aparat değişimi olmayan tercih edilir
    """
    EPS=0.5
    asgn={p:{} for p in ORDERS}
    seqs=make_empty_seqs()
    for p in sorted(ORDERS,key=lambda x:DUE[x]):
        for tau in TAU:
            best_s,best_r,best_co=None,float('inf'),True
            for s in OPT[tau]:
                t=_seq_end(seqs[s],tau,s)
                co=bool(seqs[s]) and FIX[seqs[s][-1]][tau]!=FIX[p][tau]
                if co: t+=SU[s]
                fin=t+PROC[p][tau][s]
                est=fin if s==INT[tau] else fin+tt(s,'depot')
                # Birincil: est < best_r - ε  → kesin geçiş
                # İkincil: ε içinde, aparat değişimi yok  → tercih
                if est<best_r-EPS or (abs(est-best_r)<=EPS and co<best_co):
                    best_r,best_s,best_co=est,s,co
            asgn[p][tau]=best_s; seqs[best_s].append(p)
    return asgn,seqs,build_routes(needed_firms(asgn))

# ═══════════════════════════════════════════════════════════════
#  YOK ETME OPERATÖRLERİ
#  [D1] Her operatör tam olarak q adet (p,tau) çifti kaldırır
# ═══════════════════════════════════════════════════════════════
def _remove(asgn,seqs,p,tau):
    s=asgn[p].pop(tau); seqs[s].remove(p)

def destroy_random(asgn,seqs,routes,q,rng):
    a,s,r=copy_sol(asgn,seqs,routes)
    parts=[(p,tau) for p in ORDERS for tau in TAU]
    removed=rng.sample(parts,min(q,len(parts)))
    for p,tau in removed: _remove(a,s,p,tau)
    return a,s,build_routes(needed_firms_partial(a)),removed

def destroy_worst_tardiness(asgn,seqs,routes,q,rng):
    """
    [D1] Her (p,tau) çifti T[p] değeriyle puanlanır.
    En yüksek puanlı tam q çift kaldırılır.
    """
    _,_,_,T=evaluate(asgn,seqs,routes)
    cands=sorted([(T[p],p,tau) for p in ORDERS for tau in TAU],reverse=True)
    a,s,r=copy_sol(asgn,seqs,routes)
    removed=[]
    for _,p,tau in cands:
        if len(removed)>=q: break
        _remove(a,s,p,tau); removed.append((p,tau))
    return a,s,build_routes(needed_firms_partial(a)),removed

def destroy_bottleneck(asgn,seqs,routes,q,rng):
    """
    [D1] Her (p,tau) çifti r_arr[p][tau] ile puanlanır.
    Depoya en geç ulaşan tam q parça kaldırılır.
    """
    _,r_arr,_,_=evaluate(asgn,seqs,routes)
    cands=sorted([(r_arr[p][tau],p,tau) for p in ORDERS for tau in TAU],reverse=True)
    a,s,r=copy_sol(asgn,seqs,routes)
    removed=[]
    for _,p,tau in cands:
        if len(removed)>=q: break
        _remove(a,s,p,tau); removed.append((p,tau))
    return a,s,build_routes(needed_firms_partial(a)),removed

# ═══════════════════════════════════════════════════════════════
#  ONARMA OPERATÖRLERİ
#  [D3] _try_insert her denemede rotayı NN ile sıfırdan kurar.
#       Bu, raporda tanımlanan "insertion" mantığının genişletilmiş
#       halidir: opsiyon+konum seçiminde rota maliyeti de dahil edilir.
# ═══════════════════════════════════════════════════════════════
def _try_insert(asgn,seqs,p,tau,s,k):
    seqs[s].insert(k,p)
    asgn.setdefault(p,{})[tau]=s
    routes=build_routes(needed_firms_partial(asgn))
    cost=obj(asgn,seqs,routes)
    seqs[s].pop(k); asgn[p].pop(tau)
    return cost

def _best_insertion(asgn,seqs,p,tau):
    best=(float('inf'),None,None)
    for s in OPT[tau]:
        for k in range(len(seqs[s])+1):
            c=_try_insert(asgn,seqs,p,tau,s,k)
            if c<best[0]: best=(c,s,k)
    return best

def repair_greedy(asgn,seqs,routes,removed):
    a,s,r=copy_sol(asgn,seqs,routes)
    for p,tau in sorted(removed,key=lambda x:DUE[x[0]]):
        _,bs,bk=_best_insertion(a,s,p,tau)
        s[bs].insert(bk,p); a.setdefault(p,{})[tau]=bs
    return a,s,build_routes(needed_firms(a))

def repair_regret(asgn,seqs,routes,removed):
    a,s,r=copy_sol(asgn,seqs,routes)
    pending=list(removed)
    while pending:
        best_reg,best_ch=-1.,None
        for p,tau in pending:
            costs=sorted((_try_insert(a,s,p,tau,os,k),os,k)
                         for os in OPT[tau] for k in range(len(s[os])+1))
            reg=(costs[1][0]-costs[0][0]) if len(costs)>1 else 0.
            if reg>best_reg:
                best_reg=reg; best_ch=(p,tau,costs[0][1],costs[0][2])
        p,tau,bs,bk=best_ch
        s[bs].insert(bk,p); a.setdefault(p,{})[tau]=bs
        pending.remove((p,tau))
    return a,s,build_routes(needed_firms(a))

# ═══════════════════════════════════════════════════════════════
#  ALNS ANA DÖNGÜSÜ
# ═══════════════════════════════════════════════════════════════
def alns(n_iter=600,q=4,T0=6.,alpha=.97,rho=.20,
         d1=3.,d2=2.,d3=1.,seed=42,verbose=True):
    rng=random.Random(seed)
    DEST=[destroy_random,destroy_worst_tardiness,destroy_bottleneck]
    REP =[repair_greedy,repair_regret]
    wd=[1.]*len(DEST); wr=[1.]*len(REP)
    d_calls=[0]*len(DEST); r_calls=[0]*len(REP)
    d_impr=[0.]*len(DEST);  r_impr=[0.]*len(REP)

    a,s,r=build_initial()
    a_b,s_b,r_b=copy_sol(a,s,r)
    f_curr=obj(a,s,r); f_best=f_curr; f_init=f_curr
    Theta=T0; hist=[f_best]

    if verbose:
        print(f"  Başlangıç: obj={f_init:.4f}")
        print(f"  {'Iter':>5}  {'Best':>8}  {'Curr':>8}  {'Temp':>9}")

    for it in range(1,n_iter+1):
        di=rng.choices(range(len(DEST)),weights=wd)[0]
        ri=rng.choices(range(len(REP)), weights=wr)[0]
        d_calls[di]+=1; r_calls[ri]+=1

        a2,s2,r2,rem=DEST[di](a,s,r,q,rng)
        a2,s2,r2=REP[ri](a2,s2,r2,rem)
        f_new=obj(a2,s2,r2)

        if f_new<f_best:
            a_b,s_b,r_b=copy_sol(a2,s2,r2); f_best=f_new; delta=d1
            d_impr[di]+=f_curr-f_new; r_impr[ri]+=f_curr-f_new
        elif f_new<f_curr: delta=d2
        else: delta=d3

        if f_new<f_curr or rng.random()<math.exp(max(-500,-(f_new-f_curr)/Theta)):
            a,s,r=a2,s2,r2; f_curr=f_new

        wd[di]=(1-rho)*wd[di]+rho*delta
        wr[ri]=(1-rho)*wr[ri]+rho*delta
        Theta*=alpha; hist.append(f_best)

        if verbose and it%100==0:
            print(f"  {it:5d}  {f_best:8.4f}  {f_curr:8.4f}  {Theta:9.5f}")

    op_stats={'destroy_names':['Random','WorstTard','Bottleneck'],
              'repair_names': ['Greedy','Regret'],
              'd_calls':d_calls,'r_calls':r_calls,
              'd_weights':wd,'r_weights':wr,
              'd_impr':d_impr,'r_impr':r_impr}
    return a_b,s_b,r_b,f_best,f_init,hist,op_stats

# ═══════════════════════════════════════════════════════════════
#  MAIN (M01 referans instance)
# ═══════════════════════════════════════════════════════════════
def main():
    print("="*60)
    print("  ARP-SA v2 — ALNS Çözücü (Düzeltilmiş)")
    print("="*60)
    t0=time.time()
    a,s,r,fb,fi,hist,op=alns(n_iter=600,q=4,seed=42,verbose=True)
    cpu=time.time()-t0
    _,r_arr,C,T=evaluate(a,s,r)
    print(f"\n  En iyi: {fb:.4f}  |  Başlangıç: {fi:.4f}  |  İyileştirme: {100*(fi-fb)/max(fi,1e-9):.1f}%  |  CPU: {cpu:.2f}s")
    print(f"\n  {'Sipariş':<8} {'Due':>5} {'C':>8} {'T':>7}  Atama")
    for p in ORDERS:
        print(f"  {p:<8} {DUE[p]:>5} {C[p]:>8.2f} {T[p]:>7.2f}  "+" ".join(f"{tau}→{a[p][tau]}" for tau in TAU))
    out=os.path.join(os.path.dirname(os.path.abspath(__file__)),'arp_sa_v2_results.json')
    with open(out,'w') as f2:
        json.dump({'f_best':round(fb,4),'f_init':round(fi,4),'cpu':round(cpu,3),
                   'assignments':{p:dict(a[p]) for p in ORDERS},
                   'completion':{p:round(C[p],3) for p in ORDERS},
                   'tardiness':{p:round(T[p],3) for p in ORDERS},
                   'routes':r,'convergence':hist},f2,indent=2)
    print(f"\n  Kaydedildi → {out}")

if __name__=='__main__': main()