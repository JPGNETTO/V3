import { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.jpgnetto.carteiraproventos',
  appName: 'Carteira Proventos',
  webDir: 'dist',
  server: {
    androidScheme: 'https',
    // Permite que o app acesse servidores HTTP (não-https), como o AI Bridge
    // no seu PC via Tailscale (http://100.100.195.84:4000). Sem isto, o Android
    // bloqueia a conexão por segurança e o app fica "offline".
    cleartext: true
  },
  android: {
    backgroundColor: '#0a0a14',
    // Reforça a permissão de tráfego HTTP em todo o app
    allowMixedContent: true
  }
};

export default config;
