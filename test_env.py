try:
    import selenium
    import webdriver_manager
    print("✅ 工具箱已就緒")
except ImportError:
    print("❌ 缺少套件")
    print("請複製並執行以下指令來安裝：")
    print("pip install selenium webdriver-manager")
