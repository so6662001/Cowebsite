<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">申请试用</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          免费体验钢贸宝ERP、库准MES等产品，开启您的钢铁贸易数字化之旅
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
          <span itemprop="name" class="text-gray-900 font-medium">申请试用</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Trial Form -->
    <section class="py-16 lg:py-24 bg-white">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid lg:grid-cols-5 gap-16">
          <div class="lg:col-span-3">
            <h2 class="text-2xl font-bold text-gray-900 mb-2">填写试用申请</h2>
            <p class="text-steel-light mb-8">请填写以下信息，我们的专业顾问将在24小时内与您联系，为您安排产品演示和试用。</p>

            <form @submit.prevent="submitTrial" class="space-y-6">
              <div class="grid sm:grid-cols-2 gap-4">
                <div>
                  <label for="trial-name" class="block text-sm font-medium text-gray-700 mb-1">姓名 *</label>
                  <input id="trial-name" v-model="form.contactName" type="text" required
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="您的姓名">
                </div>
                <div>
                  <label for="trial-phone" class="block text-sm font-medium text-gray-700 mb-1">手机号 *</label>
                  <input id="trial-phone" v-model="form.phone" type="tel" required pattern="[0-9]{11}"
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="11位手机号码">
                </div>
              </div>
              <div>
                <label for="trial-company" class="block text-sm font-medium text-gray-700 mb-1">公司名称 *</label>
                <input id="trial-company" v-model="form.companyName" type="text" required
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all" placeholder="公司全称">
              </div>
              <div class="grid sm:grid-cols-2 gap-4">
                <div>
                  <label for="trial-type" class="block text-sm font-medium text-gray-700 mb-1">企业类型 *</label>
                  <select id="trial-type" v-model="form.companyType" required
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all">
                    <option value="">请选择</option>
                    <option value="steel_trade">钢铁贸易企业</option>
                    <option value="steel_manufacture">钢铁生产加工企业</option>
                    <option value="other">其他</option>
                  </select>
                </div>
                <div>
                  <label for="trial-scale" class="block text-sm font-medium text-gray-700 mb-1">企业规模</label>
                  <select id="trial-scale" v-model="form.companyScale"
                    class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all">
                    <option value="">请选择</option>
                    <option value="1-10">1-10人</option>
                    <option value="10-50">10-50人</option>
                    <option value="50-200">50-200人</option>
                    <option value="200+">200人以上</option>
                  </select>
                </div>
              </div>
              <div>
                <label class="block text-sm font-medium text-gray-700 mb-3">感兴趣的产品 *</label>
                <div class="grid sm:grid-cols-2 gap-3">
                  <label v-for="product in productOptions" :key="product.value"
                    class="flex items-center gap-3 p-3 rounded-xl border cursor-pointer transition-colors"
                    :class="form.products.includes(product.value) ? 'border-primary bg-primary-light' : 'border-gray-200 hover:border-gray-300'">
                    <input type="checkbox" :value="product.value" v-model="form.products" class="sr-only">
                    <div class="w-5 h-5 rounded border-2 flex items-center justify-center flex-shrink-0"
                      :class="form.products.includes(product.value) ? 'border-primary bg-primary' : 'border-gray-300'">
                      <svg v-if="form.products.includes(product.value)" class="w-3 h-3 text-white" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                      </svg>
                    </div>
                    <span class="text-sm text-gray-700">{{ product.label }}</span>
                  </label>
                </div>
              </div>
              <div>
                <label for="trial-remark" class="block text-sm font-medium text-gray-700 mb-1">补充说明</label>
                <textarea id="trial-remark" v-model="form.remark" rows="3"
                  class="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all resize-none" placeholder="请描述您的业务需求或关注的问题..."></textarea>
              </div>
              <button type="submit" :disabled="submitting || form.products.length === 0"
                class="w-full px-6 py-4 bg-primary text-white font-semibold rounded-xl hover:bg-primary-dark transition-colors disabled:opacity-50 disabled:cursor-not-allowed text-lg">
                {{ submitting ? '提交中...' : '提交试用申请' }}
              </button>
              <p v-if="submitSuccess" class="text-green-600 text-center font-medium">
                申请提交成功！我们的顾问将在24小时内与您联系。
              </p>
            </form>
          </div>

          <div class="lg:col-span-2">
            <div class="sticky top-24 space-y-8">
              <div class="bg-bg-section rounded-2xl p-8">
                <h3 class="text-lg font-semibold text-gray-900 mb-6">为什么选择普讯？</h3>
                <div class="space-y-4">
                  <div v-for="benefit in benefits" :key="benefit.title" class="flex items-start gap-3">
                    <svg class="w-5 h-5 text-primary flex-shrink-0 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                      <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                    </svg>
                    <div>
                      <h4 class="font-medium text-gray-900 text-sm">{{ benefit.title }}</h4>
                      <p class="text-xs text-steel-light mt-1">{{ benefit.desc }}</p>
                    </div>
                  </div>
                </div>
              </div>
              <div class="bg-primary rounded-2xl p-8 text-white">
                <h3 class="text-lg font-semibold mb-4">需要立即咨询？</h3>
                <p class="text-white/80 text-sm mb-6">拨打服务中心电话或添加微信，我们的专业顾问随时为您解答。</p>
                <div class="space-y-3">
                  <div class="flex items-center gap-3">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"/>
                    </svg>
                    <span class="text-sm">025-xxxxxxxx</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
import { reactive, ref } from 'vue'

const form = reactive({
  contactName: '',
  phone: '',
  companyName: '',
  companyType: '',
  companyScale: '',
  products: [],
  remark: ''
})

const submitting = ref(false)
const submitSuccess = ref(false)

const productOptions = [
  { value: 'gangmaobao_erp', label: '钢贸宝 ERP' },
  { value: 'gangqitong_erp', label: '钢企通 ERP' },
  { value: 'kuzhun_mes', label: '库准 生产MES' },
  { value: 'smart_finance', label: '智能财务' },
  { value: 'huodaizi', label: '货袋子平台' },
  { value: 'ai_worker', label: 'AI数字员工' }
]

const benefits = [
  { title: '免费试用体验', desc: '无需付费即可体验产品核心功能' },
  { title: '专业顾问1对1', desc: '资深行业顾问为您提供个性化建议' },
  { title: '快速上线部署', desc: '专业团队支持快速部署和数据迁移' },
  { title: '全国服务覆盖', desc: '覆盖全国的服务团队，随时提供支持' },
  { title: '以周为单位迭代', desc: '产品快速迭代，持续满足新需求' }
]

const submitTrial = async () => {
  submitting.value = true
  try {
    const response = await fetch('/api/trial', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ...form,
        products: form.products.join(',')
      })
    })
    if (response.ok) {
      submitSuccess.value = true
      Object.assign(form, {
        contactName: '', phone: '', companyName: '',
        companyType: '', companyScale: '', products: [], remark: ''
      })
      setTimeout(() => { submitSuccess.value = false }, 8000)
    }
  } catch (e) {
    alert('提交失败，请稍后重试或直接拨打电话 025-xxxxxxxx 联系我们。')
  } finally {
    submitting.value = false
  }
}
</script>
