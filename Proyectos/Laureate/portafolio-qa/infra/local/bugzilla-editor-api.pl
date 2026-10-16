use strict;
use warnings;
use lib '/var/www/html/bugzilla';
chdir '/var/www/html/bugzilla' or die $!;
require Bugzilla;
require Bugzilla::Constants;
require Bugzilla::User;
require Bugzilla::User::APIKey;
require JSON::XS;
Bugzilla->usage_mode(Bugzilla::Constants::USAGE_MODE_CMDLINE());
my $admin = Bugzilla::User->new({ name => $ENV{BUGZILLA_ADMIN_EMAIL} });
die 'Administrador no encontrado' unless $admin;
Bugzilla->set_user($admin);
my $login = 'portafolio-editor-api@portafolio.local';
my $user = Bugzilla::User->new({ name => $login });
if (!$user) {
    open my $random, '<', '/dev/urandom' or die $!;
    read($random, my $bytes, 24) == 24 or die 'Entropy unavailable';
    $user = Bugzilla::User->create({ login_name => $login, realname => 'Portafolio QA API lectura y edicion', cryptpassword => unpack('H*', $bytes) });
}
$user->set_groups({ add => ['editbugs', 'canconfirm'] });
$user->update();
$user = Bugzilla::User->new($user->id);
die 'Faltan permisos de edicion' unless $user->in_group('editbugs') && $user->in_group('canconfirm');
my $key = Bugzilla::User::APIKey->create({ user_id => $user->id, description => 'Portafolio QA lectura y edicion 2026-10-16' });
print JSON::XS::encode_json({ api_key => $key->api_key, login => $login, editbugs => JSON::XS::true(), canconfirm => JSON::XS::true() });

