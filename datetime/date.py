import datetime

def extract_year_month(val):
    if isinstance(val, datetime):
        return val.strftime("%Y-%m")
    if isinstance(val, str) and val:
        val = val.strip()
        try:
            for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y-%m", "%Y/%m"]:
                try:
                    return datetime.strptime(val, fmt).strftime("%Y-%m")
                except:
                    pass
        except:
            pass
    return None
