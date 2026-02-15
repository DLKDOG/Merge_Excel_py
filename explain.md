graph TD
    %% 스타일 정의
    classDef configClass fill:#f9f,stroke:#333,stroke-width:2px;
    classDef uiClass fill:#bbf,stroke:#333,stroke-width:2px;
    classDef logicClass fill:#dfd,stroke:#333,stroke-width:2px;
    classDef queueShape fill:#ff9,stroke:#f66,stroke-width:2px,stroke-dasharray: 5 5;

    %% 1. 진입점
    subgraph MainEntry ["1. 프로그램 시작 (Main)"]
        main("if __name__ == '__main__'")
    end

    %% 2. 설정 클래스
    subgraph ConfigLayer ["2. 설정 (AppConfig)"]
        direction TB
        conf_init("AppConfig 초기화")
        conf_setup("setup_logging")
    end

    %% 3. UI 클래스
    subgraph UILayer ["4. 화면 (MergeApp)"]
        direction TB
        ui_init("__init__")
        ui_init_ui("_init_ui")
        ui_start("start_process (버튼 클릭)")
        ui_thread("Thread 생성 및 시작")
        ui_worker("_run_worker (스레드 실행)")
        ui_poll("_poll_queue (화면 갱신)")
    end

    %% 4. 로직 클래스 (중첩 제거하여 안정성 확보)
    subgraph LogicLayer ["3. 로직 (ExcelProcessor)"]
        direction TB
        logic_init("__init__")
        logic_process("process (메인 작업)")

        %% 내부 함수들
        merge("_merge_files")
        sort("_sort_dataframe")
        chart_ex("_create_excel_chart")
        chart_pl("_create_plotly_html")
    end

    %% 5. 통신 (큐)
    queue_bucket(("Queue (데이터 통로)")):::queueShape

    %% --- 연결 관계 ---
    main --> conf_init
    conf_init --> conf_setup
    main --> logic_init
    main --> ui_init
    ui_init --> ui_init_ui

    %% 버튼 클릭 흐름
    ui_start --> ui_thread
    ui_thread -.-> ui_worker

    %% 작업 위임 (UI -> Logic)
    ui_worker == "1. 작업 지시 (호출)" ==> logic_process

    %% 로직 내부 흐름
    logic_process --> merge
    logic_process --> sort
    logic_process --> chart_ex
    logic_process --> chart_pl

    %% 보고 체계
    logic_process -- "2. 진행률 보고 (put)" --> queue_bucket
    ui_poll -- "3. 진행률 확인 (get)" --> queue_bucket

    %% 스타일 적용
    class conf_init,conf_setup configClass;
    class ui_init,ui_init_ui,ui_start,ui_thread,ui_worker,ui_poll uiClass;
    class logic_init,logic_process,merge,sort,chart_ex,chart_pl logicClass;
