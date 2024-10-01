plugins {
    kotlin("jvm") version "2.0.20"
}

group = "mrvin"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

val apachePoi = "5.3.0"

dependencies {
    implementation("org.apache.poi:poi-ooxml:$apachePoi")
    testImplementation(kotlin("test"))
}

tasks.test { useJUnitPlatform() }

kotlin { jvmToolchain(21) }