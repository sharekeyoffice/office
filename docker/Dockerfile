# Dockerfile — nginx serving the prebuilt dist/.
#
# Build: docker build -t sharekey-office:local -f docker/Dockerfile .
# Run:   docker run --rm -p 8080:80 sharekey-office:local
#
# Production: tag with the wrapper version (e.g. v9.3.0+0) and push
# to the cluster's image registry.

FROM nginx:1.27-alpine

# Replace the default config with ours (sets COOP/COEP, gzip_static, no SPA fallback).
RUN rm /etc/nginx/conf.d/default.conf
COPY docker/nginx.conf /etc/nginx/nginx.conf

# Copy the built static site.
# `dist/` must exist at build time — run `npm run build` (bash build/build.sh) first.
COPY dist /usr/share/nginx/html

RUN chmod -R a+rX /usr/share/nginx/html

# Healthcheck: ensure /edit.html is served.
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s \
  CMD wget --quiet --tries=1 --spider http://localhost/edit.html || exit 1

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
