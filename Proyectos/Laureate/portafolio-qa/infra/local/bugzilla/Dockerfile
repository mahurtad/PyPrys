FROM debian:12-slim

ARG BUGZILLA_REPO=https://github.com/bugzilla/bugzilla.git
ARG BUGZILLA_BRANCH=release-5.2-stable

ENV DEBIAN_FRONTEND=noninteractive
ENV BUGZILLA_HOME=/var/www/html/bugzilla

# Runtime packages for Apache CGI, Perl, MariaDB connectivity and CPAN builds.
# libsoap-lite-perl, libxml-parser-perl y libxmlrpc-lite-perl habilitan xmlrpc.cgi,
# la API que consume la integracion de Kiwi TCMS.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        apache2 \
        build-essential \
        ca-certificates \
        cpanminus \
        default-libmysqlclient-dev \
        git \
        libcgi-pm-perl \
        libdatetime-perl \
        libdbd-mariadb-perl \
        libdbd-mysql-perl \
        libdbi-perl \
        libemail-mime-perl \
        libemail-sender-perl \
        libemail-address-xs-perl \
        libjson-xs-perl \
        liblist-moreutils-perl \
        libmath-random-isaac-perl \
        libmoo-perl \
        libsoap-lite-perl \
        libtemplate-perl \
        libtimedate-perl \
        libxml-parser-perl \
        libxmlrpc-lite-perl \
        liburi-perl \
        libdbix-connector-perl \
        mariadb-client \
        netcat-openbsd \
        patch \
        perl \
        wget \
    && rm -rf /var/lib/apt/lists/*

# Pull Bugzilla from the official repository and branch.
RUN git clone --depth 1 --branch "${BUGZILLA_BRANCH}" "${BUGZILLA_REPO}" "${BUGZILLA_HOME}" \
    && chown -R www-data:www-data "${BUGZILLA_HOME}"

WORKDIR ${BUGZILLA_HOME}

# Install Bugzilla 5.2 required modules at versions accepted by checksetup.pl.
RUN cpanm --notest \
        CGI \
        DateTime \
        DateTime::TimeZone \
        DBI \
        DBIx::Connector \
        Moo \
        Template \
        Email::Sender \
        Email::Address::XS \
        Email::MIME \
        URI \
        List::MoreUtils \
        Math::Random::ISAAC \
        JSON::XS \
        JSON::RPC \
        Test::Taint \
        TimeDate \
    && rm -rf /root/.cpanm /tmp/*

COPY apache-bugzilla.conf /etc/apache2/sites-available/bugzilla.conf
COPY entrypoint.sh /usr/local/bin/bugzilla-entrypoint.sh
COPY localconfig.template /usr/local/share/bugzilla/localconfig.template

RUN sed -i 's/\r$//' /usr/local/bin/bugzilla-entrypoint.sh \
    && chmod +x /usr/local/bin/bugzilla-entrypoint.sh \
    && a2enmod cgi headers rewrite expires \
    && a2dissite 000-default.conf \
    && a2ensite bugzilla.conf \
    && mkdir -p "${BUGZILLA_HOME}/localconfig.d" "${BUGZILLA_HOME}/data/attachments" \
    && chown -R www-data:www-data "${BUGZILLA_HOME}/data" "${BUGZILLA_HOME}/localconfig.d"

EXPOSE 8080

ENTRYPOINT ["/usr/local/bin/bugzilla-entrypoint.sh"]
CMD ["apache2ctl", "-D", "FOREGROUND"]
