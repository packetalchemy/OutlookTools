# ساخت نصب‌کننده VSTO با Inno Setup

## پیش‌نیاز:
1. **Inno Setup** رو دانلود و نصب کنید:
   ```
   https://jrsoftware.org/isinfo.php
   ```
   (رایگان و سبک — ~5MB)

## مراحل ساخت Setup.exe:

### مرحله ۱: پروژه رو Build کنید
```
VS 2022 → OutlookTools.sln → Build → Release
```
خروجی در: `OutlookTools/bin/Release/`

### مرحله ۲: Inno Setup Compiler رو باز کنید
```
Start Menu → Inno Setup Compiler
```

### مرحله ۳: فایل .iss رو باز کنید
```
File → Open
→ OutlookTools/Setup/OutlookToolsSetup.iss
```

### مرحله ۴: Compile کنید
```
Build → Compile
یا
Ctrl+F9
```

### مرحله ۵: فایل خروجی
```
Setup/installer_output/OutlookTools_Setup_1.2.0.exe
```

## ویژگی‌های نصب‌کننده:
- ✅ ثبت خودکار COM Add-in
- ✅ بستن اوت‌لوک قبل از نصب
- ✅ Uninstall با حذف کامل
- ✅ چک کردن .NET Framework
- ✅ منوی Start
- ✅ حذف خودکار فایل‌های موقت

## خروجی نهایی:
```
Setup/installer_output/OutlookTools_Setup_1.2.0.exe (~2MB)
```

این فایل رو به هر کسی بدید تا نصب کنه!
