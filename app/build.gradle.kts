import com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar

plugins {
    java
    id("com.github.johnrengelman.shadow") version "8.1.1"
}

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("org.apache.commons:commons-csv:1.10.0")
    implementation("com.github.mygreen:excel-cellformatter:0.12")
    implementation("org.slf4j:slf4j-simple:2.0.5")
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
