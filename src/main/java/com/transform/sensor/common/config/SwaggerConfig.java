package com.transform.sensor.common.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

import org.springframework.web.servlet.config.annotation.EnableWebMvc;
import springfox.documentation.builders.ApiInfoBuilder;
import springfox.documentation.builders.RequestHandlerSelectors;
import springfox.documentation.spi.DocumentationType;
import springfox.documentation.spring.web.plugins.Docket;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

@Configuration
@EnableSwagger2
@ComponentScan(basePackages = {"com.hiekn.rest"}) //必须存在 扫描的API Controller package name 也可以直接扫描class (basePackageClasses)
public class SwaggerConfig {

	@Bean
	public Docket robot() {
		return new Docket(DocumentationType.SWAGGER_2).groupName("For File Transform").select()
				.apis(RequestHandlerSelectors.basePackage("com.hiekn.rest")).build()
				.apiInfo(new ApiInfoBuilder().title("For File Transform").description("transform")
						.version("0.0.1").build());
	}
}
