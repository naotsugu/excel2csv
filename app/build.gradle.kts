import com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar

plugins {
    java
    id("com.github.johnrengelman.shadow") version "6.1.0"
}

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.0.0")
    implementation("org.apache.poi:poi-ooxml:5.0.0")
    implementation("org.apache.commons:commons-csv:1.8")
    implementation("com.github.mygreen:excel-cellformatter:0.12")
    implementation("org.slf4j:slf4j-simple:1.7.30")
}

tasks {
    named<ShadowJar>("shadowJar") {
        archiveBaseName.set("excel2csv")
        archiveClassifier.set("")
        archiveVersion.set("")
        manifest {
            attributes(mapOf("Main-Class" to "com.mammb.excel2csv.Main"))
        }
    }
}

tasks {
    build {
        dependsOn(shadowJar)
    }
}
