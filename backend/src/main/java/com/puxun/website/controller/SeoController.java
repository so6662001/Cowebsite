package com.puxun.website.controller;

import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class SeoController {

    @GetMapping(value = "/sitemap.xml", produces = MediaType.APPLICATION_XML_VALUE)
    public ResponseEntity<String> sitemap() {
        String baseUrl = "https://www.puxunsoft.com";
        String sitemap = """
            <?xml version="1.0" encoding="UTF-8"?>
            <urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
              <url>
                <loc>%s/</loc>
                <changefreq>weekly</changefreq>
                <priority>1.0</priority>
              </url>
              <url>
                <loc>%s/about</loc>
                <changefreq>monthly</changefreq>
                <priority>0.8</priority>
              </url>
              <url>
                <loc>%s/products</loc>
                <changefreq>weekly</changefreq>
                <priority>0.9</priority>
              </url>
              <url>
                <loc>%s/solutions</loc>
                <changefreq>weekly</changefreq>
                <priority>0.9</priority>
              </url>
              <url>
                <loc>%s/partners</loc>
                <changefreq>monthly</changefreq>
                <priority>0.7</priority>
              </url>
              <url>
                <loc>%s/contact</loc>
                <changefreq>monthly</changefreq>
                <priority>0.7</priority>
              </url>
              <url>
                <loc>%s/trial</loc>
                <changefreq>monthly</changefreq>
                <priority>0.8</priority>
              </url>
              <url>
                <loc>%s/learning</loc>
                <changefreq>monthly</changefreq>
                <priority>0.6</priority>
              </url>
            </urlset>
            """.formatted(baseUrl, baseUrl, baseUrl, baseUrl, baseUrl, baseUrl, baseUrl, baseUrl);
        return ResponseEntity.ok(sitemap);
    }

    @GetMapping(value = "/robots.txt", produces = MediaType.TEXT_PLAIN_VALUE)
    public ResponseEntity<String> robots() {
        String robots = """
            User-agent: *
            Allow: /
            Disallow: /api/

            Sitemap: https://www.puxunsoft.com/sitemap.xml
            """;
        return ResponseEntity.ok(robots);
    }
}
