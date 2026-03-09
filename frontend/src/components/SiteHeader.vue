<template>
  <header class="fixed top-0 left-0 right-0 z-50 transition-all duration-300"
    :class="scrolled ? 'bg-white/95 backdrop-blur-md shadow-md' : 'bg-transparent'">
    <nav class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8" aria-label="主导航">
      <div class="flex items-center justify-between h-16 lg:h-20">
        <router-link to="/" class="flex items-center gap-3" aria-label="普讯软件首页">
          <div class="w-10 h-10 rounded-lg bg-primary flex items-center justify-center">
            <span class="text-white font-bold text-xl">普</span>
          </div>
          <div>
            <span class="text-lg font-bold" :class="scrolled ? 'text-gray-900' : 'text-white'">普讯软件</span>
            <span class="hidden sm:block text-xs" :class="scrolled ? 'text-gray-500' : 'text-white/70'">钢铁贸易数字化解决方案</span>
          </div>
        </router-link>

        <div class="hidden lg:flex items-center gap-1">
          <router-link v-for="item in navItems" :key="item.path" :to="item.path"
            class="px-4 py-2 rounded-lg text-sm font-medium transition-colors"
            :class="isActive(item.path)
              ? 'bg-primary text-white'
              : scrolled ? 'text-gray-700 hover:text-primary hover:bg-primary-light' : 'text-white/90 hover:text-white hover:bg-white/10'">
            {{ item.name }}
          </router-link>
          <a href="https://www.huodaizi.com" target="_blank" rel="noopener"
            class="px-4 py-2 rounded-lg text-sm font-medium transition-colors"
            :class="scrolled ? 'text-gray-700 hover:text-primary hover:bg-primary-light' : 'text-white/90 hover:text-white hover:bg-white/10'">
            找钢材
          </a>
        </div>

        <div class="hidden lg:flex items-center gap-3">
          <router-link to="/trial"
            class="px-5 py-2.5 bg-primary text-white text-sm font-medium rounded-lg hover:bg-primary-dark transition-colors shadow-sm">
            申请试用
          </router-link>
        </div>

        <button @click="mobileMenuOpen = !mobileMenuOpen" class="lg:hidden p-2 rounded-lg"
          :class="scrolled ? 'text-gray-700' : 'text-white'" aria-label="切换菜单">
          <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path v-if="!mobileMenuOpen" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h16" />
            <path v-else stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>

      <transition enter-active-class="transition duration-200 ease-out" enter-from-class="opacity-0 -translate-y-2"
        enter-to-class="opacity-100 translate-y-0" leave-active-class="transition duration-150 ease-in"
        leave-from-class="opacity-100 translate-y-0" leave-to-class="opacity-0 -translate-y-2">
        <div v-if="mobileMenuOpen" class="lg:hidden pb-4">
          <div class="bg-white rounded-xl shadow-xl p-4 space-y-1">
            <router-link v-for="item in navItems" :key="item.path" :to="item.path" @click="mobileMenuOpen = false"
              class="block px-4 py-3 rounded-lg text-sm font-medium transition-colors"
              :class="isActive(item.path) ? 'bg-primary text-white' : 'text-gray-700 hover:bg-gray-100'">
              {{ item.name }}
            </router-link>
            <a href="https://www.huodaizi.com" target="_blank" rel="noopener"
              class="block px-4 py-3 rounded-lg text-sm font-medium text-gray-700 hover:bg-gray-100">
              找钢材
            </a>
            <router-link to="/trial" @click="mobileMenuOpen = false"
              class="block px-4 py-3 rounded-lg text-sm font-medium text-center bg-primary text-white hover:bg-primary-dark">
              申请试用
            </router-link>
          </div>
        </div>
      </transition>
    </nav>
  </header>
</template>

<script setup>
import { ref, onMounted, onUnmounted } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const scrolled = ref(false)
const mobileMenuOpen = ref(false)

const navItems = [
  { name: '首页', path: '/' },
  { name: '关于我们', path: '/about' },
  { name: '产品中心', path: '/products' },
  { name: '解决方案', path: '/solutions' },
  { name: '合作企业', path: '/partners' },
  { name: '学习中心', path: '/learning' },
  { name: '联系我们', path: '/contact' }
]

const isActive = (path) => route.path === path

const handleScroll = () => {
  scrolled.value = window.scrollY > 50
}

onMounted(() => {
  window.addEventListener('scroll', handleScroll)
  handleScroll()
})

onUnmounted(() => {
  window.removeEventListener('scroll', handleScroll)
})
</script>
