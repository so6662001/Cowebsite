<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">联系我们</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          无论是购前咨询、商务合作还是售后服务，我们随时为您提供支持
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
          <span itemprop="name" class="text-gray-900 font-medium">联系我们</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Contact Info -->
    <section class="py-16 lg:py-24 bg-white" aria-labelledby="contact-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid lg:grid-cols-3 gap-8 mb-16">
          <div v-for="office in offices" :key="office.title"
            class="bg-bg-section rounded-2xl p-8 card-hover">
            <div class="w-12 h-12 rounded-xl bg-primary-light flex items-center justify-center mb-4">
              <svg class="w-6 h-6 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" :d="office.icon" />
              </svg>
            </div>
            <h3 class="text-lg font-semibold text-gray-900 mb-2">{{ office.title }}</h3>
            <p class="text-steel-light text-sm">{{ office.address }}</p>
          </div>
        </div>

        <div class="grid lg:grid-cols-2 gap-16">
          <div>
            <h2 id="contact-title" class="text-2xl font-bold text-gray-900 mb-8">联系方式</h2>
            <div class="space-y-6">
              <div class="flex items-start gap-4">
                <div class="w-10 h-10 rounded-lg bg-primary-light flex items-center justify-center flex-shrink-0">
                  <svg class="w-5 h-5 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"/>
                  </svg>
                </div>
                <div>
                  <h3 class="font-semibold text-gray-900">服务中心电话</h3>
                  <p class="text-steel-light">025-xxxxxxxx</p>
                </div>
              </div>
              <div class="flex items-start gap-4">
                <div class="w-10 h-10 rounded-lg bg-primary-light flex items-center justify-center flex-shrink-0">
                  <svg class="w-5 h-5 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 12h.01M12 12h.01M16 12h.01M21 12c0 4.418-4.03 8-9 8a9.863 9.863 0 01-4.255-.949L3 20l1.395-3.72C3.512 15.042 3 13.574 3 12c0-4.418 4.03-8 9-8s9 3.582 9 8z"/>
                  </svg>
                </div>
                <div>
                  <h3 class="font-semibold text-gray-900">微信咨询</h3>
                  <p class="text-steel-light">扫描购前咨询微信二维码</p>
                </div>
              </div>
              <div class="flex items-start gap-4">
                <div class="w-10 h-10 rounded-lg bg-primary-light flex items-center justify-center flex-shrink-0">
                  <svg class="w-5 h-5 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 10l4.553-2.276A1 1 0 0121 8.618v6.764a1 1 0 01-1.447.894L15 14M5 18h8a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"/>
                  </svg>
                </div>
                <div>
                  <h3 class="font-semibold text-gray-900">视频号 / 抖音</h3>
                  <p class="text-steel-light">关注我们获取行业资讯和产品动态</p>
                </div>
              </div>
            </div>
          </div>

          <div>
            <h2 class="text-2xl font-bold text-gray-900 mb-8">商务合作</h2>
            <p class="text-steel-light mb-6">
              如果您对商务合作感兴趣，请填写以下信息，我们将通过企业微信尽快与您取得联系。
            </p>
            <form @submit.prevent="submitContact" class="space-y-4">
              <div class="grid sm:grid-cols-2 gap-4">
                <div>
                  <label for="contact-name" class="block text-sm font-medium text-gray-700 mb-1">姓名 *</label>
                  <input id="contact-name" v-model="form.name" type="text" required
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="您的姓名">
                </div>
                <div>
                  <label for="contact-phone" class="block text-sm font-medium text-gray-700 mb-1">电话 *</label>
                  <input id="contact-phone" v-model="form.phone" type="tel" required
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="联系电话">
                </div>
              </div>
              <div>
                <label for="contact-company" class="block text-sm font-medium text-gray-700 mb-1">公司名称 *</label>
                <input id="contact-company" v-model="form.company" type="text" required
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="公司名称">
              </div>
              <div>
                <label for="contact-message" class="block text-sm font-medium text-gray-700 mb-1">合作意向</label>
                <textarea id="contact-message" v-model="form.message" rows="4"
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all resize-none" placeholder="请描述您的合作意向..."></textarea>
              </div>
              <button type="submit" :disabled="submitting"
                class="w-full px-6 py-3 bg-primary text-white font-medium rounded-xl hover:bg-primary-dark transition-colors disabled:opacity-50 disabled:cursor-not-allowed">
                {{ submitting ? '提交中...' : '提交合作意向' }}
              </button>
              <p v-if="submitSuccess" class="text-green-600 text-sm text-center">提交成功！我们将尽快与您联系。</p>
            </form>
          </div>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
import { reactive, ref } from 'vue'
import { useHead } from '@unhead/vue'

useHead({
  script: [
    {
      type: 'application/ld+json',
      innerHTML: JSON.stringify({
        '@context': 'https://schema.org',
        '@type': 'LocalBusiness',
        name: '南京普讯管理软件有限公司',
        description: '钢铁贸易数字化解决方案提供商',
        address: [
          { '@type': 'PostalAddress', addressLocality: '上海', addressCountry: 'CN', description: '公司总部' },
          { '@type': 'PostalAddress', addressLocality: '南京', addressCountry: 'CN', description: '研发中心 / 销售中心' }
        ],
        telephone: '025-xxxxxxxx'
      })
    }
  ]
})

const offices = [
  {
    title: '总部地址',
    address: '上海市',
    icon: 'M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4'
  },
  {
    title: '研发中心',
    address: '南京市',
    icon: 'M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4'
  },
  {
    title: '销售中心',
    address: '南京市',
    icon: 'M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z'
  }
]

const form = reactive({
  name: '',
  phone: '',
  company: '',
  message: ''
})

const submitting = ref(false)
const submitSuccess = ref(false)

const submitContact = async () => {
  submitting.value = true
  try {
    const response = await fetch('/api/contact', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(form)
    })
    if (response.ok) {
      submitSuccess.value = true
      Object.assign(form, { name: '', phone: '', company: '', message: '' })
      setTimeout(() => { submitSuccess.value = false }, 5000)
    }
  } catch (e) {
    alert('提交失败，请稍后重试或直接拨打电话联系我们。')
  } finally {
    submitting.value = false
  }
}
</script>
