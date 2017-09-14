#!groovy
pipeline {
	agent {
		label 'windows'
	}
	stages {
		stage('Build') {
			steps {
				bat 'pyinstaller --onefile --windowed --name wcexport wcexport.py'
				bat 'pyinstaller --onefile --windowed --name docexport docexport.py'
			}
		}
		stage('Commit Binaries') {
			steps {
				bat 'copy dist\\wcexport.exe .'
				bat 'copy dist\\docexport.exe .'
				bat 'git commit *.exe -m "Updated Windows Executables" || exit 0 && git push origin HEAD:master'
			}
		}
	}
}
