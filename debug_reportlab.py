try:
    import reportlab
    print(f"ReportLab Version: {reportlab.__version__}")
    print(f"File: {reportlab.__file__}")

    import reportlab.pdfbase.ttfonts
    print("dir(reportlab.pdfbase.ttfonts):")
    print(dir(reportlab.pdfbase.ttfonts))

    from reportlab.pdfbase import ttfonts
    print(f"TTFFont in ttfonts? {'TTFFont' in dir(ttfonts)}")

except Exception as e:
    print(f"Error: {e}")
