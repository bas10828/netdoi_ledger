def tlv(s, depth=0):
    out=[]
    i=0
    while i+4<=len(s):
        tag=s[i:i+2]; ln=int(s[i+2:i+4]); val=s[i+4:i+4+ln]
        out.append((tag,ln,val)); i+=4+ln
    return out

samples = {
 "77456 (ttb paybill, ref 260618103442386876)":"00390006000001010301102182606181034423868765102TH9104ECF7",
 "77494":"004600060000010103002022520260618170921230099571085102TH9104ED08",
 "77524":"004600060000010103002022520260619100542240008637085102TH910416BA",
}
for name,q in samples.items():
    print(f"=== {name} ===")
    for tag,ln,val in tlv(q):
        print(f"  {tag} (len {ln}): {val}")
        if tag=="00":
            for t2,l2,v2 in tlv(val):
                print(f"      sub {t2} (len {l2}): {v2}")
