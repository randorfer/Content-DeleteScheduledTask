def buildfile

node ("ubuntu1404") {
    // Mark the code checkout 'stage'....
    stage 'Checkout'

    // checkout the location of this Jenkinsfile
    deleteDir()
    checkout scm

    stage 'Setup Tools Environment'
    tool name: 'Ant 1.9.7', type: 'hudson.tasks.Ant$AntInstallation'
   
    def antHome = tool 'Ant 1.9.7'

    echo 'Ant Home = ' + antHome

  	echo 'Copying Content Build Tools Artifacts'
    step([$class: 'CopyArtifact', filter: '**/*.zip', fingerprintArtifacts: true, flatten: true, projectName: 'Tanium/Content-BuildTools/master', selector: [$class: 'StatusBuildSelector', stable: false], target: 'tools'])

    if(isUnix()) {
   		sh antHome+ "/bin/ant unzip-tools"
   	}
   	else {
		bat antHome+ "/bin/ant unzip-tools"
   	}
	
	stage 'Transfer to Common Jenkinsfile'
	echo 'Loading common build groovy file'
	buildfile = load 'tools/TaniumContentBuildTools/jenkinsbuild.groovy'
	
	echo 'Initiating buildAll in common file'
	buildfile.buildAll()
   
}
