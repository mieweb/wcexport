#!groovy
pipeline {
	agent {
		label 'windows'
	}
	stages {
		stage('Build') {
			bat 'pyinstaller --onefile --windowed --name wcexport wcexport.py'
			bat 'pyinstaller --onefile --windowed --name docexport docexport.py'
			bat 
		}
		stage('Commit Binaries') {
			bat 'copy dist\\wcexport.exe .'
			bat 'copy dist\\docexport.exe .'
			bat 'git commit *.exe -m "Updated Windows Executables"'
			bat 'git push origin HEAD:master'
		}
	}
}
