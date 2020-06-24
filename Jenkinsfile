#!/usr/bin/env groovy

pipeline {
  agent { 
    label 'wollmux'
  }
	
  options {
    disableConcurrentBuilds()
  }

  tools {
    jdk 'Java11'
  }
	
  stages {
    stage('Build') {
      steps {
        withMaven(
          maven: 'mvn',
          mavenLocalRepo: '.repo',
          mavenSettingsConfig: 'org.jenkinsci.plugins.configfiles.maven.MavenSettingsConfig1441715654272',
          publisherStrategy: 'EXPLICIT') {
          sh "mvn -DdryRun clean package"
        }
      }
    }
    stage('Quality Gate') {
      steps {
        script {
          if (GIT_BRANCH == 'master') {
            withMaven(
              maven: 'mvn',
              mavenLocalRepo: '.repo',
              mavenSettingsConfig: 'org.jenkinsci.plugins.configfiles.maven.MavenSettingsConfig1441715654272',
              publisherStrategy: 'EXPLICIT') {
              withSonarQubeEnv('SonarQube') {
                sh "mvn $SONAR_MAVEN_GOAL \
                -Dsonar.host.url=$SONAR_HOST_URL \
                -Dsonar.branch.name=${GIT_BRANCH}"
              }
            }
          } else {
            withMaven(
              maven: 'mvn',
              mavenLocalRepo: '.repo',
              mavenSettingsConfig: 'org.jenkinsci.plugins.configfiles.maven.MavenSettingsConfig1441715654272',
              publisherStrategy: 'EXPLICIT') {
              withSonarQubeEnv('SonarQube') {
                sh "mvn $SONAR_MAVEN_GOAL \
                  -Dsonar.host.url=$SONAR_HOST_URL \
                  -Dsonar.branch.name=${GIT_BRANCH} \
                  -Dsonar.branch.target=${env.CHANGE_TARGET}"
              }
            }
          }
        }
      }
    }
    stage('Artifactory Deploy') {
      when {
        branch "master"
      }
      steps {
        withMaven(
          maven: 'mvn',
          mavenLocalRepo: '.repo',
          mavenSettingsConfig: 'org.jenkinsci.plugins.configfiles.maven.MavenSettingsConfig1441715654272',
          publisherStrategy: 'EXPLICIT') {
          script {
			      def server = Artifactory.server('-122848432@1441782548261')
			      def rtMaven = Artifactory.newMavenBuild()
			      rtMaven.resolver server: server, releaseRepo: 'libs-release', snapshotRepo: 'libs-snapshot'
			      rtMaven.deployer server: server, releaseRepo: 'libs-release-local', snapshotRepo: 'libs-snapshot-local'
			      def buildInfo = rtMaven.run pom: 'pom.xml', goals: 'install'
			      server.publishBuildInfo buildInfo
          }
        }
      }
    }
  }
}
