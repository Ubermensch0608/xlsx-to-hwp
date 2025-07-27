def run_with_interrupt(func):
    try:
        while True:
            print("작업 실행 중... 종료하려면 Ctrl + C")
            func()
    except KeyboardInterrupt:
        print("\n[사용자에 의해 종료됨]")
