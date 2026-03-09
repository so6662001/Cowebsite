<template>
  <div>
    <!-- Hero Section -->
    <section class="gradient-hero relative overflow-hidden min-h-[90vh] flex items-center" aria-label="首页横幅">
      <div class="absolute inset-0 opacity-10">
        <div class="absolute top-20 left-10 w-72 h-72 bg-primary rounded-full blur-3xl"></div>
        <div class="absolute bottom-20 right-10 w-96 h-96 bg-secondary rounded-full blur-3xl"></div>
      </div>
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-24 pb-16 relative z-10">
        <div class="grid lg:grid-cols-2 gap-12 items-center">
          <div class="fade-in">
            <div class="inline-flex items-center gap-2 px-4 py-2 bg-white/10 backdrop-blur rounded-full text-white/90 text-sm mb-8">
              <span class="w-2 h-2 bg-green-400 rounded-full animate-pulse"></span>
              深耕钢贸行业近20年 · 服务近3000家企业
            </div>
            <h1 class="text-4xl sm:text-5xl lg:text-6xl font-bold text-white leading-tight mb-6">
              钢铁贸易<br>
              <span class="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-cyan-300">数字化解决方案</span><br>
              提供商
            </h1>
            <p class="text-lg text-white/80 leading-relaxed mb-8 max-w-xl">
              普讯软件提供钢贸宝ERP、钢企通ERP、库准MES、智能财务等全套钢铁行业数智化系统，以及货袋子钢铁产业互联网交易平台。我们给企业的不是一套软件，而是一套业务发展的经营系统。
            </p>
            <div class="flex flex-wrap gap-4">
              <router-link to="/trial"
                class="px-8 py-4 bg-white text-primary font-semibold rounded-xl hover:bg-gray-100 transition-all shadow-lg hover:shadow-xl">
                免费申请试用
              </router-link>
              <router-link to="/solutions"
                class="px-8 py-4 bg-white/10 backdrop-blur text-white font-semibold rounded-xl border border-white/20 hover:bg-white/20 transition-all">
                了解解决方案
              </router-link>
            </div>
          </div>
          <div class="hidden lg:block">
            <div class="relative">
              <div class="bg-white/10 backdrop-blur-lg rounded-2xl p-8 border border-white/10">
                <div class="grid grid-cols-2 gap-4">
                  <div v-for="stat in heroStats" :key="stat.label" class="bg-white/10 rounded-xl p-5 text-center">
                    <div class="text-3xl font-bold text-white mb-1">{{ stat.value }}</div>
                    <div class="text-sm text-white/70">{{ stat.label }}</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- Stats Bar (Mobile) -->
    <section class="lg:hidden bg-primary text-white py-8" aria-label="核心数据">
      <div class="max-w-7xl mx-auto px-4 grid grid-cols-2 gap-6">
        <div v-for="stat in heroStats" :key="stat.label" class="text-center">
          <div class="text-2xl font-bold">{{ stat.value }}</div>
          <div class="text-sm text-white/80">{{ stat.label }}</div>
        </div>
      </div>
    </section>

    <!-- Products Overview -->
    <section class="py-20 lg:py-28 gradient-section" aria-labelledby="products-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <h2 id="products-title" class="text-3xl lg:text-4xl font-bold text-gray-900 mb-4">核心产品矩阵</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            全栈自主研发，为钢铁贸易与生产制造企业提供全套经营数智化系统
          </p>
        </div>
        <div class="grid sm:grid-cols-2 lg:grid-cols-3 gap-8">
          <article v-for="product in products" :key="product.name"
            class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100 card-hover">
            <div class="w-14 h-14 rounded-xl flex items-center justify-center mb-6" :class="product.bgColor">
              <span class="text-2xl">{{ product.icon }}</span>
            </div>
            <h3 class="text-xl font-bold text-gray-900 mb-3">{{ product.name }}</h3>
            <p class="text-steel-light leading-relaxed mb-4">{{ product.desc }}</p>
            <div class="flex flex-wrap gap-2">
              <span v-for="tag in product.tags" :key="tag"
                class="px-3 py-1 bg-gray-100 text-gray-600 text-xs rounded-full">{{ tag }}</span>
            </div>
          </article>
        </div>
        <div class="text-center mt-12">
          <router-link to="/products" class="inline-flex items-center gap-2 text-primary font-medium hover:underline">
            查看全部产品详情
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/></svg>
          </router-link>
        </div>
      </div>
    </section>

    <!-- Why Choose Us -->
    <section class="py-20 lg:py-28 bg-white" aria-labelledby="why-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <h2 id="why-title" class="text-3xl lg:text-4xl font-bold text-gray-900 mb-4">为什么选择普讯？</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            我们给企业的不是一套软件，而是一套业务发展的经营系统
          </p>
        </div>
        <div class="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
          <div v-for="advantage in advantages" :key="advantage.title"
            class="flex gap-4 p-6 rounded-2xl hover:bg-gray-50 transition-colors">
            <div class="w-12 h-12 rounded-xl bg-primary-light flex items-center justify-center flex-shrink-0">
              <svg class="w-6 h-6 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" :d="advantage.icon" />
              </svg>
            </div>
            <div>
              <h3 class="text-lg font-semibold text-gray-900 mb-2">{{ advantage.title }}</h3>
              <p class="text-sm text-steel-light leading-relaxed">{{ advantage.desc }}</p>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- Solutions Preview -->
    <section class="py-20 lg:py-28 bg-bg-section" aria-labelledby="solutions-title">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <h2 id="solutions-title" class="text-3xl lg:text-4xl font-bold text-gray-900 mb-4">行业解决方案</h2>
          <p class="text-lg text-steel-light max-w-2xl mx-auto">
            针对钢贸企业和钢铁生产制造企业，提供从业务增长到降本增效的全方位数智化方案
          </p>
        </div>
        <div class="grid lg:grid-cols-2 gap-8">
          <div class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100">
            <h3 class="text-xl font-bold text-gray-900 mb-6 flex items-center gap-3">
              <span class="w-10 h-10 rounded-lg bg-blue-100 flex items-center justify-center text-blue-600 text-lg">贸</span>
              钢贸企业解决方案
            </h3>
            <div class="grid sm:grid-cols-2 gap-3">
              <div v-for="s in tradeSolutions" :key="s" class="flex items-center gap-2 p-3 rounded-lg bg-gray-50">
                <svg class="w-4 h-4 text-primary flex-shrink-0" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/></svg>
                <span class="text-sm text-gray-700">{{ s }}</span>
              </div>
            </div>
          </div>
          <div class="bg-white rounded-2xl p-8 shadow-sm border border-gray-100">
            <h3 class="text-xl font-bold text-gray-900 mb-6 flex items-center gap-3">
              <span class="w-10 h-10 rounded-lg bg-green-100 flex items-center justify-center text-green-600 text-lg">产</span>
              生产制造企业解决方案
            </h3>
            <div class="grid sm:grid-cols-2 gap-3">
              <div v-for="s in mfgSolutions" :key="s" class="flex items-center gap-2 p-3 rounded-lg bg-gray-50">
                <svg class="w-4 h-4 text-green-600 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/></svg>
                <span class="text-sm text-gray-700">{{ s }}</span>
              </div>
            </div>
          </div>
        </div>
        <div class="text-center mt-10">
          <router-link to="/solutions"
            class="inline-flex items-center gap-2 px-8 py-3 bg-primary text-white font-medium rounded-xl hover:bg-primary-dark transition-colors">
            查看全部解决方案
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/></svg>
          </router-link>
        </div>
      </div>
    </section>

    <!-- FAQ for AI SEO -->
    <section class="py-20 lg:py-28 bg-white" aria-labelledby="faq-title">
      <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="text-center mb-16">
          <h2 id="faq-title" class="text-3xl lg:text-4xl font-bold text-gray-900 mb-4">常见问题</h2>
          <p class="text-lg text-steel-light">了解更多关于普讯和钢铁行业数字化的信息</p>
        </div>
        <div class="space-y-4">
          <details v-for="faq in faqs" :key="faq.q" class="group bg-gray-50 rounded-xl overflow-hidden">
            <summary class="flex items-center justify-between p-6 cursor-pointer hover:bg-gray-100 transition-colors">
              <h3 class="text-base font-semibold text-gray-900 pr-4">{{ faq.q }}</h3>
              <svg class="w-5 h-5 text-gray-400 group-open:rotate-180 transition-transform flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"/></svg>
            </summary>
            <div class="px-6 pb-6 text-steel-light leading-relaxed">{{ faq.a }}</div>
          </details>
        </div>
      </div>
    </section>

    <!-- CTA -->
    <section class="py-20 gradient-hero" aria-label="开始使用">
      <div class="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h2 class="text-3xl lg:text-4xl font-bold text-white mb-6">开启钢铁贸易数字化转型之旅</h2>
        <p class="text-lg text-white/80 mb-10 max-w-2xl mx-auto">
          帮助更多的企业实现更高效率、更有质量的发展。产品升级以周为单位，快速满足市场和客户需求。
        </p>
        <div class="flex flex-wrap justify-center gap-4">
          <router-link to="/trial" class="px-8 py-4 bg-white text-primary font-semibold rounded-xl hover:bg-gray-100 transition-all shadow-lg">
            立即申请试用
          </router-link>
          <router-link to="/contact" class="px-8 py-4 border-2 border-white text-white font-semibold rounded-xl hover:bg-white/10 transition-all">
            联系我们
          </router-link>
        </div>
      </div>
    </section>
  </div>
</template>

<script setup>
import { useHead } from '@unhead/vue'

useHead({
  script: [
    {
      type: 'application/ld+json',
      innerHTML: JSON.stringify({
        '@context': 'https://schema.org',
        '@type': 'FAQPage',
        mainEntity: [
          {
            '@type': 'Question',
            name: '普讯软件是做什么的？',
            acceptedAnswer: { '@type': 'Answer', text: '普讯软件是一家钢铁贸易数字化解决方案提供商，同时也是货袋子钢铁交易平台运营商。提供钢贸宝ERP、钢企通ERP、库准MES、智能财务等全套钢铁行业数智化系统。' }
          },
          {
            '@type': 'Question',
            name: '钢贸宝ERP适合什么样的企业使用？',
            acceptedAnswer: { '@type': 'Answer', text: '钢贸宝ERP适合所有钢铁贸易企业使用，从小型贸易商到大型集团钢贸企业，已有近3000家钢贸企业合作使用。系统支持多分公司管理、销售外出管理、财务自动化等功能。' }
          },
          {
            '@type': 'Question',
            name: '普讯软件服务了多少家企业？',
            acceptedAnswer: { '@type': 'Answer', text: '截至目前，普讯合作钢贸宝ERP企业近3000家，日活用户10000+人。合作库准仓库管理系统100+家，合作生产MES系统10+家。货袋子日均线上库存10万吨。' }
          },
          {
            '@type': 'Question',
            name: '货袋子钢铁交易平台是什么？',
            acceptedAnswer: { '@type': 'Answer', text: '货袋子是普讯运营的钢铁产业互联网交易赋能平台，日均线上库存10万吨，帮助钢铁贸易企业实现线上卖货和采购。访问 huodaizi.com 了解更多。' }
          },
          {
            '@type': 'Question',
            name: '如何申请试用普讯的产品？',
            acceptedAnswer: { '@type': 'Answer', text: '您可以通过官网的"申请试用"页面提交试用申请，填写企业信息后，我们的服务团队将尽快与您联系，为您提供产品演示和试用体验。' }
          }
        ]
      })
    }
  ]
})

const heroStats = [
  { value: '近20年', label: '深耕钢贸行业' },
  { value: '3000+', label: '合作钢贸企业' },
  { value: '10000+', label: '日活跃用户' },
  { value: '5000万', label: '累计研发投入' }
]

const products = [
  {
    name: '钢贸宝 ERP',
    icon: '📊',
    desc: '专为钢铁贸易企业打造的全流程经营数智化管理系统，覆盖采购、销售、库存、财务全链条。',
    tags: ['采购管理', '销售管理', '库存管理', '数据分析'],
    bgColor: 'bg-blue-100'
  },
  {
    name: '钢企通 ERP',
    icon: '🏭',
    desc: '面向钢铁生产制造企业的综合管理平台，整合生产、仓储、销售、财务一体化管理。',
    tags: ['生产管理', '一体化', '集团管控', '智能分析'],
    bgColor: 'bg-green-100'
  },
  {
    name: '库准 生产MES',
    icon: '📦',
    desc: '仓库与生产制造执行系统，实现仓库作业数字化管理，支持无人值守等先进场景。',
    tags: ['仓库管理', '生产执行', '无人值守', '实时监控'],
    bgColor: 'bg-purple-100'
  },
  {
    name: '智能财务',
    icon: '💰',
    desc: '钢贸行业专属智能财务系统，自动化核算、对账、结算，大幅提升财务工作效率。',
    tags: ['自动对账', '智能核算', '报表分析', '多维度'],
    bgColor: 'bg-amber-100'
  },
  {
    name: '货袋子赋能平台',
    icon: '🌐',
    desc: '钢铁产业互联网交易赋能平台，日均线上库存10万吨，助力企业线上卖货、获取客户。',
    tags: ['线上交易', '库存共享', '获客引流', '日均10万吨'],
    bgColor: 'bg-cyan-100'
  },
  {
    name: 'AI数字员工',
    icon: '🤖',
    desc: '基于AI技术的企业数字员工解决方案，智能处理重复性工作，提升运营自动化水平。',
    tags: ['人工智能', '自动化', '智能客服', '数据分析'],
    bgColor: 'bg-rose-100'
  }
]

const advantages = [
  {
    title: '近20年行业深耕',
    desc: '深入钢贸企业实战运营4年，深刻理解行业痛点和需求，解决方案更贴合实际。',
    icon: 'M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z'
  },
  {
    title: '5000万研发投入',
    desc: '累计投入研发近5000万，全栈自主研发，产品升级以周为单位，快速响应市场需求。',
    icon: 'M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z'
  },
  {
    title: '3000+企业信赖',
    desc: '合作钢贸宝ERP企业近3000家，日活用户10000+人，覆盖各类钢铁贸易企业。',
    icon: 'M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z'
  },
  {
    title: '全国服务覆盖',
    desc: '员工近100人，服务团队覆盖全国各地，确保每一位客户都能获得及时专业的服务支持。',
    icon: 'M3.055 11H5a2 2 0 012 2v1a2 2 0 002 2 2 2 0 012 2v2.945M8 3.935V5.5A2.5 2.5 0 0010.5 8h.5a2 2 0 012 2 2 2 0 104 0 2 2 0 012-2h1.064M15 20.488V18a2 2 0 012-2h3.064'
  },
  {
    title: '不只是软件',
    desc: '我们给企业的不是一套软件，而是一套业务发展的经营系统，包含经营思路、理念及落地方法。',
    icon: 'M13 10V3L4 14h7v7l9-11h-7z'
  },
  {
    title: '赋能平台理念',
    desc: '我们不是一家软件公司，而是一个赋能平台，让更多企业享受更好的服务和发展。',
    icon: 'M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.888a1 1 0 00-1.176 0l-3.976 2.888c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.784-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z'
  }
]

const tradeSolutions = [
  '业务增长解决方案',
  '降本增效解决方案',
  '企业扩张解决方案',
  '财务自动化解决方案',
  '集团分公司管理方案',
  '客户/销售外出管理方案',
  '运营自动化解决方案',
  '仓库管理解决方案',
  'AI数字员工解决方案'
]

const mfgSolutions = [
  '生产MES解决方案',
  '生产成本管理方案',
  '生产设备管理方案',
  '生产效能管理方案',
  '排产自动化解决方案',
  '生产质量管控方案',
  '工人工资自动化核算方案',
  '无人值守解决方案'
]

const faqs = [
  {
    q: '普讯软件是做什么的？',
    a: '普讯软件是一家钢铁贸易数字化解决方案提供商，同时也是货袋子钢铁交易平台运营商。我们提供钢贸宝ERP、钢企通ERP、库准MES、智能财务等全套钢铁行业数智化系统，帮助钢贸企业和钢铁生产制造企业实现数字化转型升级。公司深耕钢贸行业近20年，累计投入研发近5000万。'
  },
  {
    q: '钢贸宝ERP适合什么样的企业使用？',
    a: '钢贸宝ERP适合所有钢铁贸易企业使用，从小型贸易商到大型集团钢贸企业均可覆盖。目前已有近3000家钢贸企业合作使用。系统支持多分公司管理、销售外出管理、财务自动化、客户管理等功能，帮助企业实现全流程数字化经营。'
  },
  {
    q: '普讯软件服务了多少家企业？',
    a: '截至目前，普讯合作钢贸宝ERP企业近3000家，日活用户10000+人。合作库准仓库管理系统100+家，合作生产MES系统10+家。货袋子日均线上库存10万吨。服务团队覆盖全国各地，确保每位客户都能获得专业及时的服务。'
  },
  {
    q: '货袋子钢铁交易平台是什么？',
    a: '货袋子是普讯运营的钢铁产业互联网交易赋能平台，日均线上库存10万吨。平台帮助钢铁贸易企业实现线上卖货和采购，是钢铁行业的产业互联网平台。您可以访问 huodaizi.com 了解更多详情。'
  },
  {
    q: '如何申请试用普讯的产品？',
    a: '您可以通过官网的"申请试用"页面提交试用申请，填写企业基本信息后，我们的服务团队将在第一时间与您联系，为您提供产品演示和试用体验。我们的服务热线：025-xxxxxxxx，也可以通过微信联系我们。'
  }
]
</script>
