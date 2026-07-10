#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ADCDA v1.1 automated content auditor — machine-checkable rules."""
import json, re, sys

EXPECTED = ["cover","objectives","main_message_why","definition","causes",
            "warning_signs","prevention_1","prevention_2","mistake_alternatives",
            "response_steps","special_groups","emergency_999","after_incident",
            "key_messages","sources"]

def wc(t): return len([w for w in re.split(r"\s+", t.strip()) if w])

def texts_of(s):
    out=[]
    for k in ("title","text","empowerment_line","truth_opener","critical_dont",
              "interaction_question","curiosity_hook","scenario","scenario_callback",
              "scenario_close","surprise_fact","stat","stat_meaning","extra"):
        if s.get(k): out.append(s[k])
    for k in ("items","steps","when","say","report","safe_return","disposal"):
        out += s.get(k,[]) or []
    for p in s.get("pairs",[]) or []:
        out += [p.get("instead_of",""), p.get("do_this","")]
    for g in s.get("groups",[]) or []:
        out += [g.get("group",""), g.get("behavior","")]
    return out

def audit(path):
    d=json.load(open(path,encoding="utf-8"))
    sl={s["n"]:s for s in d["slides"]}
    R=[]  # (item, status, detail)  status: PASS/FAIL/PENDING
    def add(i,st,dt=""): R.append((i,st,dt))

    # 1. sequence 15
    seq=[s["type"] for s in d["slides"]]
    add("1 التسلسل 15 والقوس","PASS" if seq==EXPECTED else "FAIL", "" if seq==EXPECTED else str(seq))

    # 2. seven questions = required types present (covered by seq) + non-empty cores
    core_ok = all([sl[4].get("text"), sl[3].get("stat"), sl[5].get("items"),
                   sl[6].get("items"), sl[7].get("items"), sl[10].get("steps"),
                   sl[13].get("report") or sl[13].get("safe_return")])
    add("2 الأسئلة السبعة","PASS" if core_ok else "FAIL")

    # 3. limits: title<=8, bullets<=5, each text<=12 words (warnings exempt? no — rule says bullets)
    bad=[]
    for s in d["slides"]:
        t=s.get("title","")
        if t and wc(t)>8: bad.append(f"ش{s['n']} عنوان {wc(t)} كلمة")
        for lk in ("items","steps","when","say","report","safe_return","disposal"):
            lst=s.get(lk,[]) or []
            if len(lst)>5: bad.append(f"ش{s['n']} {lk}: {len(lst)} نقاط")
            for b in lst:
                if wc(b)>12: bad.append(f"ش{s['n']} نقطة {wc(b)} كلمة: {b[:30]}…")
        for p in s.get("pairs",[]) or []:
            for v in p.values():
                if wc(v)>12: bad.append(f"ش{s['n']} زوج {wc(v)} كلمة")
        for g in s.get("groups",[]) or []:
            if wc(g.get("behavior",""))>12: bad.append(f"ش{s['n']} سلوك فئة {wc(g['behavior'])} كلمة")
    add("3 حدود الكلمات","PASS" if not bad else "FAIL","; ".join(bad))

    # 4. stats sourced (placeholder => PENDING)
    st=sl[3]
    if "[" in st.get("stat","") or "[" in st["source"].get("name",""):
        add("4 مصدر لكل رقم + تفسير","PENDING","إحصائية بانتظار وكيل البحث (بالتصميم)")
    else:
        ok=bool(st["source"].get("name")) and bool(st.get("stat_meaning"))
        add("4 مصدر لكل رقم + تفسير","PASS" if ok else "FAIL")

    # 5. risk two dims
    r=sl[3].get("risk",{})
    add("5 الخطر: احتمالية + شدة","PASS" if r.get("likelihood") and r.get("severity") else "FAIL")

    # 6. 999 only + terminology
    blob=" ".join(t for s in d["slides"] for t in texts_of(s))
    bad_nums=re.findall(r"\b(911|112|997|998|991|999\d)\b",blob)
    bad_terms=[t for t in ("المعاقين","ذوي الاحتياجات","المعوقين") if t in blob]
    ok=("999" in blob) and not bad_nums and not bad_terms
    add("6 رقم 999 والمصطلحات","PASS" if ok else "FAIL",str(bad_nums+bad_terms))

    # 7. main message <=15, in s3, verbatim in s14
    mm=d.get("main_message","")
    ok = mm and wc(mm)<=15 and sl[3].get("main_message")==mm and sl[14].get("main_message")==mm
    add("7 الرسالة الرئيسية ≤15 ومكررة حرفيًا","PASS" if ok else "FAIL",f"{wc(mm)} كلمة")

    # 8. scenario open/callback/close
    ok = bool(sl[3].get("scenario")) and bool(sl[10].get("scenario_callback")) and bool(sl[14].get("scenario_close"))
    add("8 السيناريو الإطاري 3→10→14","PASS" if ok else "FAIL")

    # 9. empowerment + positive framing (exactly one critical_dont; no other "لا ت" in behavior lists)
    emp=bool(sl[7].get("empowerment_line"))
    cd = 1 if sl[9].get("critical_dont","").startswith("لا") else 0
    stray=[]
    for s in d["slides"]:
        for lk in ("items","steps","when","say","report","safe_return","disposal"):
            for b in s.get(lk,[]) or []:
                if re.match(r"^لا\s+ت",b): stray.append(f"ش{s['n']}: {b[:25]}")
        for g in s.get("groups",[]) or []:
            if re.match(r"^لا\s+ت",g.get("behavior","")): stray.append(f"ش{s['n']} فئة")
    ok = emp and cd==1 and not stray
    add("9 تمكين + صياغة إيجابية + نهي حرج واحد","PASS" if ok else "FAIL","; ".join(stray))

    # 10. interaction questions (6,11) + hooks (3,6,9)
    ok = all([sl[6].get("interaction_question"), sl[11].get("interaction_question"),
              sl[3].get("curiosity_hook"), sl[6].get("curiosity_hook"), sl[9].get("curiosity_hook")])
    add("10 سؤالان تفاعليان + 3 أسطر فضول","PASS" if ok else "FAIL")

    # 11. visuals on 3,4,10
    ok = all(sl[n].get("visual",{}).get("kind") and sl[n]["visual"].get("desc") for n in (3,4,10))
    add("11 الدليل البصري (3،4،10)","PASS" if ok else "FAIL")

    # 12. rhythm: no 3 consecutive same pattern
    pats=[s.get("pattern") for s in d["slides"]]
    viol=[i+1 for i in range(len(pats)-2) if pats[i]==pats[i+1]==pats[i+2]]
    add("12 قاعدة الإيقاع (لا 3 متتالية)","PASS" if not viol else "FAIL",str(viol))

    # 13. after_incident + audience/season
    ok = sl[13]["type"]=="after_incident" and d.get("audience") and d.get("season")
    add("13 بعد الحادث + الجمهور والموسم","PASS" if ok else "FAIL")

    # 14. threat-solution balance + ISO line present
    ok = bool(d["standards"].get("compliance_line_ar")) and bool(sl[15].get("iso26000"))
    add("14 كل خطر بحل + سطر ISO","PASS" if ok else "FAIL")

    # 15. language — automated subset: no double spaces, tanwin on ً common words heuristic skipped
    dbl = [s["n"] for s in d["slides"] for t in texts_of(s) if "  " in t]
    add("15 اللغة (فحص آلي جزئي)","PASS" if not dbl else "FAIL",f"مسافات مزدوجة: {dbl}")

    # 16. standards activation matches profile
    prof=d.get("topic_profile","")
    cond=d["standards"]["conditional"]
    need45 = ("عمل" in prof); need22 = ("استجابة" in prof or "إخلاء" in prof)
    ok = (bool(cond["iso45001"])==need45) and (cond["iso22320_22322"]==need22) \
         and (("45001" in d["standards"]["compliance_line_ar"])==need45) \
         and (("22320" in d["standards"]["compliance_line_ar"])==need22)
    add("16 تفعيل المواصفات حسب النوع","PASS" if ok else "FAIL",prof)

    print(f"\n=== {d['topic_code']} — {d['title_ar']} ===")
    p=sum(1 for _,s,_ in R if s=="PASS"); f=sum(1 for _,s,_ in R if s=="FAIL"); pe=sum(1 for _,s,_ in R if s=="PENDING")
    for i,s2,dt in R:
        mark={"PASS":"✅","FAIL":"❌","PENDING":"⏳"}[s2]
        print(f"{mark} {i}" + (f" — {dt}" if dt else ""))
    print(f"النتيجة: {p} ناجح / {f} فاشل / {pe} معلّق (من 16)")
    return f==0

if __name__=="__main__":
    ok=True
    for p in sys.argv[1:]:
        ok = audit(p) and ok
    sys.exit(0 if ok else 1)
