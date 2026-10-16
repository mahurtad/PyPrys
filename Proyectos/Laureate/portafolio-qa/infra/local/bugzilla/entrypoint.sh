#!/usr/bin/env bash
set -euo pipefail

BUGZILLA_HOME="${BUGZILLA_HOME:-/var/www/html/bugzilla}"
LOCALCONFIG_DIR="${BUGZILLA_HOME}/localconfig.d"
LOCALCONFIG_FILE="${LOCALCONFIG_DIR}/localconfig"
LOCALCONFIG_LINK="${BUGZILLA_HOME}/localconfig"

required_env() {
  local name="$1"
  if [ -z "${!name:-}" ]; then
    echo "Missing required environment variable: ${name}" >&2
    exit 1
  fi
}

required_env MYSQL_HOST
required_env MYSQL_DATABASE
required_env MYSQL_USER
required_env MYSQL_PASSWORD
required_env BUGZILLA_ADMIN_EMAIL
required_env BUGZILLA_ADMIN_PASSWORD
required_env BUGZILLA_URL

# Apache escucha en el mismo puerto que se publica en el host para que la URL
# de Bugzilla sea identica desde el navegador y desde otros contenedores.
export BUGZILLA_LISTEN_PORT="${BUGZILLA_LISTEN_PORT:-8080}"
export BUGZILLA_SERVER_NAME="${BUGZILLA_SERVER_NAME:-bugzilla.local}"

# La imagen no incluye MTA, y con el valor por defecto (Sendmail) Bugzilla
# falla al crear bugs o comentar por API. None desactiva el envio de correo;
# use SMTP + BUGZILLA_SMTP_SERVER si necesita notificaciones reales.
export BUGZILLA_MAIL_DELIVERY_METHOD="${BUGZILLA_MAIL_DELIVERY_METHOD:-None}"
export BUGZILLA_SMTP_SERVER="${BUGZILLA_SMTP_SERVER:-localhost}"
echo "Listen ${BUGZILLA_LISTEN_PORT}" > /etc/apache2/ports.conf

echo "Waiting for MariaDB at ${MYSQL_HOST}:3306..."
until mysqladmin ping -h"${MYSQL_HOST}" -u"${MYSQL_USER}" -p"${MYSQL_PASSWORD}" --silent; do
  sleep 2
done

mkdir -p "${LOCALCONFIG_DIR}" "${BUGZILLA_HOME}/data/attachments"

if [ ! -f "${LOCALCONFIG_FILE}" ]; then
  echo "Creating persistent Bugzilla localconfig..."
  site_secret="$(perl -Mbytes -e 'open my $fh, "<", "/dev/urandom"; read($fh, my $b, 64); print unpack("H*", $b)')"
  sed \
    -e "s|__MYSQL_HOST__|${MYSQL_HOST}|g" \
    -e "s|__MYSQL_DATABASE__|${MYSQL_DATABASE}|g" \
    -e "s|__MYSQL_USER__|${MYSQL_USER}|g" \
    -e "s|__MYSQL_PASSWORD__|${MYSQL_PASSWORD}|g" \
    -e "s|__SITE_WIDE_SECRET__|${site_secret}|g" \
    /usr/local/share/bugzilla/localconfig.template > "${LOCALCONFIG_FILE}"
fi

perl -0pi -e "
  s|^\s*\\\$db_driver\s*=.*?;|\\\$db_driver = 'mariadb';|m;
  s|^\s*\\\$db_host\s*=.*?;|\\\$db_host = '\$ENV{MYSQL_HOST}';|m;
  s|^\s*\\\$db_name\s*=.*?;|\\\$db_name = '\$ENV{MYSQL_DATABASE}';|m;
  s|^\s*\\\$db_user\s*=.*?;|\\\$db_user = '\$ENV{MYSQL_USER}';|m;
  s|^\s*\\\$db_pass\s*=.*?;|\\\$db_pass = '\$ENV{MYSQL_PASSWORD}';|m;
" "${LOCALCONFIG_FILE}"

ln -sfn "${LOCALCONFIG_FILE}" "${LOCALCONFIG_LINK}"
chown -R www-data:www-data "${LOCALCONFIG_DIR}" "${BUGZILLA_HOME}/data"
chmod 600 "${LOCALCONFIG_FILE}"

cd "${BUGZILLA_HOME}"

echo "Running Bugzilla setup..."
printf '%s\n%s\n%s\n%s\n' \
  "${BUGZILLA_ADMIN_EMAIL}" \
  "Bugzilla Administrator" \
  "${BUGZILLA_ADMIN_PASSWORD}" \
  "${BUGZILLA_ADMIN_PASSWORD}" | perl checksetup.pl

echo "Updating Bugzilla parameters (urlbase, mail_delivery_method)..."
perl -I. -MBugzilla -MBugzilla::Constants -MBugzilla::Config -e 'Bugzilla->usage_mode(USAGE_MODE_CMDLINE); Bugzilla::Config::SetParam("urlbase", $ENV{BUGZILLA_URL}); Bugzilla::Config::SetParam("mail_delivery_method", $ENV{BUGZILLA_MAIL_DELIVERY_METHOD}); Bugzilla::Config::SetParam("smtpserver", $ENV{BUGZILLA_SMTP_SERVER}) if $ENV{BUGZILLA_MAIL_DELIVERY_METHOD} eq "SMTP"; Bugzilla::Config::write_params();'

chown -R www-data:www-data "${BUGZILLA_HOME}/data" "${LOCALCONFIG_DIR}"

# Clean up stale Apache sockets/PID files to prevent cgid initialization failure (AH01243)
rm -f /var/run/apache2/socks/cgisock.* /var/run/apache2/apache2.pid

exec "$@"
