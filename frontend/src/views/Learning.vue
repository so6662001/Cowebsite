<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">学习中心</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          不仅是软件培训，更是一套经营企业的思路、理念以及落地方法
        </p>
      </div>
    </section>

    <!-- Breadcrumb -->
    <nav class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4" aria-label="面包屑导航">
      <ol class="flex items-center gap-2 text-sm text-gray-500" itemscope itemtype="https://schema.org/BreadcrumbList">
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <router-link to="/" itemprop="item"><span itemprop="name">首页</span></router-link>
          <meta itemprop="position" content="1" />
        </li>
        <li class="text-gray-300">/</li>
        <li itemprop="itemListElement" itemscope itemtype="https://schema.org/ListItem">
          <span itemprop="name" class="text-gray-900 font-medium">学习中心</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Login Section -->
    <section class="py-16 lg:py-24 bg-white" aria-labelledby="login-title">
      <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid lg:grid-cols-2 gap-16 items-center">
          <div>
            <h2 id="login-title" class="text-3xl font-bold text-gray-900 mb-6">普讯学习平台</h2>
            <p class="text-steel-light leading-relaxed mb-6">
              普讯学习平台为合作企业提供全方位的学习资源，包括系统操作培训、行业知识分享、企业经营管理课程等。
              我们服务企业的过程不是简单的培训，而是给你一套经营企业的思路、理念以及落地方法。
            </p>
            <div class="space-y-4 mb-8">
              <div v-for="feature in features" :key="feature.title" class="flex items-start gap-3">
                <div class="w-8 h-8 rounded-lg bg-primary-light flex items-center justify-center flex-shrink-0">
                  <svg class="w-4 h-4 text-primary" fill="currentColor" viewBox="0 0 20 20">
                    <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                  </svg>
                </div>
                <div>
                  <h3 class="font-medium text-gray-900 text-sm">{{ feature.title }}</h3>
                  <p class="text-xs text-steel-light mt-1">{{ feature.desc }}</p>
                </div>
              </div>
            </div>
          </div>

          <div class="bg-bg-section rounded-2xl p-8">
            <h3 class="text-xl font-semibold text-gray-900 mb-6 text-center">登录学习平台</h3>
            <form @submit.prevent="handleLogin" class="space-y-4">
              <div>
                <label for="learn-account" class="block text-sm font-medium text-gray-700 mb-1">账号</label>
                <input id="learn-account" v-model="loginForm.account" type="text" required
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="请输入账号">
              </div>
              <div>
                <label for="learn-password" class="block text-sm font-medium text-gray-700 mb-1">密码</label>
                <input id="learn-password" v-model="loginForm.password" type="password" required
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="请输入密码">
              </div>
              <button type="submit"
                class="w-full px-6 py-3 bg-primary text-white font-medium rounded-xl hover:bg-primary-dark transition-colors">
                登录
              </button>
              <p class="text-center text-sm text-steel-light">
                还没有账号？请联系您的服务顾问开通
              </p>
            </form>
          </div>
        </div>
      </div>
    </section>

    <!-- Learning Categories -->
    <section class="py-16 lg:py-24 bg-bg-section" aria-labelledby="categories-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <h2 id="categories-title" class="text-3xl font-bold text-gray-900 mb-12 text-center">学习内容分类</h2>
        <div class="grid sm:grid-cols-2 lg:grid-cols-3 gap-8">
          <div v-for="category in categories" :key="category.title"
            class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100 card-hover">
            <div class="text-4xl mb-4">{{ category.icon }}</div>
            <h3 class="text-lg font-semibold text-gray-900 mb-3">{{ category.title }}</h3>
            <p class="text-steel-light text-sm leading-relaxed">{{ category.desc }}</p>
          </div>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
import { reactive } from 'vue'

const loginForm = reactive({
  account: '',
  password: ''
})

const handleLogin = () => {
  alert('学习平台登录功能即将上线，请联系服务顾问了解详情。')
}

const features = [
  { title: '系统操作培训', desc: '钢贸宝ERP、库准MES等产品的详细操作教程' },
  { title: '行业知识分享', desc: '钢铁贸易行业最新动态、市场分析' },
  { title: '企业经营管理', desc: '经营理念、管理方法、实战案例分享' },
  { title: '赋能培训课程', desc: '企业赋能培训、钢神经营俱乐部' }
]

const categories = [
  { icon: '📖', title: '产品教程', desc: '钢贸宝ERP、钢企通ERP、库准MES等产品的详细使用教程，帮助您快速上手。' },
  { icon: '📊', title: '行业报告', desc: '钢铁贸易行业趋势分析、市场报告、价格行情等专业内容。' },
  { icon: '💡', title: '经营方法论', desc: '普讯多年服务钢贸企业总结的经营理念、管理方法和落地方案。' },
  { icon: '🎯', title: '实战案例', desc: '成功客户案例分享，学习其他钢贸企业数字化转型的实践经验。' },
  { icon: '🎓', title: '赋能培训', desc: '企业赋能培训课程，提升团队数字化素养和经营能力。' },
  { icon: '🏆', title: '钢神俱乐部', desc: '钢神经营俱乐部资源，与行业精英交流，共享商业机会。' }
]
</script>
