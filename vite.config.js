import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  // Replace <your-repo-name> below with the exact name of this GitHub repo,
  // e.g. if your repo is github.com/minalkashif1/hr-leave-dashboard,
  // this should be base: '/hr-leave-dashboard/'
  base: '/<vroozi-leave-analytics>/',
})
