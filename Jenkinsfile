pipeline {
    agent any

    environment {
        PYTHON_PATH = 'C:\\Users\\LKiruba\\AppData\\Local\\Programs\\Python\\Python311\\python.exe'
        PIP_PATH = 'C:\\Users\\LKiruba\\AppData\\Local\\Programs\\Python\\Python311\\Scripts\\pip.exe'
    }

    stages {
        stage('Checkout SCM') {
            steps {
                checkout scmGit(branches: [[name: '*/master']], extensions: [], userRemoteConfigs: [[url: 'https://github.com/KirubaLakshminarayanan/Configini_Jsontoxl.git']])
            }
        }
        stage('Install Dependencies') {
            steps {
                bat "${PYTHON_PATH} --version"
                bat "${PIP_PATH} install pytest pytest-playwright pytest-html"
            }
        }
        stage('Execute Python File') {
            steps {
                bat 'cd'
                bat 'dir'
                bat 'if not exist jsontoexcelconfig.py (echo jsontoexcelconfig.py not found && exit /b 1)'
                bat """
                ${PYTHON_PATH} jsontoexcelconfig.py > script_output.txt 2>&1
                type script_output.txt
                if errorlevel 1 exit /b %errorlevel%
                """
            }
        }
        stage('Run Pytest') {
            steps {
                bat 'cd'
                bat 'dir'
                bat "${PYTHON_PATH} -m pytest test_jsontoexcelconfig.py -v -s --html=playwright-report/report.html --self-contained-html"
            }
        }
        stage('Verify Report Generation') {
            steps {
                dir('playwright-report') {
                    bat 'dir'
                }
            }
        }
        stage('Publish HTML Report') {
            steps {
                publishHTML([
                    allowMissing: false,
                    alwaysLinkToLastBuild: true,
                    escapeUnderscores: false,
                    keepAll: true, 
                    reportDir: 'playwright-report', 
                    reportFiles: 'report.html', 
                    reportName: 'Pytest HTML Report', 
                    reportTitles: 'Json to Excel Converter Pytest Report'
                ])
            }
        }
    }
}
