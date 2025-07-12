# ── STAGE 1: Build with a Maven image that includes JDK 21 ───────────
FROM maven:3.9.8-eclipse-temurin-21 AS builder

WORKDIR /workspace
COPY pom.xml .
RUN mvn dependency:go-offline -B

COPY src ./src
RUN mvn clean package -DskipTests -B

# ── STAGE 2: Run on a slim JRE 21 image ────────────────────────────────
FROM eclipse-temurin:21-jre-alpine

USER appuser
WORKDIR /app
COPY --from=builder /workspace/target/registro-cleaner-0.0.1-SNAPSHOT.jar app.jar
EXPOSE 8080
ENTRYPOINT ["java","-jar","app.jar"]
