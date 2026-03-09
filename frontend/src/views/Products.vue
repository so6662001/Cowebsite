<template>
  <div>
    <!-- Hero -->
    <section class="gradient-hero pt-32 pb-20">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
        <h1 class="text-4xl lg:text-5xl font-bold text-white mb-6">产品中心</h1>
        <p class="text-xl text-white/80 max-w-3xl mx-auto">
          全栈自主研发，为钢铁贸易与生产制造企业提供全套经营数智化产品体系
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
          <span itemprop="name" class="text-gray-900 font-medium">产品中心</span>
          <meta itemprop="position" content="2" />
        </li>
      </ol>
    </nav>

    <!-- Products Detail -->
    <section v-for="(product, index) in products" :key="product.name"
      :class="index % 2 === 0 ? 'bg-white' : 'bg-bg-section'"
      class="py-16 lg:py-24" :aria-labelledby="'product-' + index">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div class="grid lg:grid-cols-2 gap-12 items-center"
          :class="index % 2 !== 0 ? 'lg:grid-flow-col-dense' : ''">
          <div :class="index % 2 !== 0 ? 'lg:col-start-2' : ''">
            <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-medium mb-4"
              :class="product.tagClass">
              {{ product.tag }}
            </div>
            <h2 :id="'product-' + index" class="text-3xl font-bold text-gray-900 mb-4">{{ product.name }}</h2>
            <p class="text-lg text-steel-light mb-6 leading-relaxed">{{ product.desc }}</p>
            <h3 class="text-base font-semibold text-gray-900 mb-4">核心功能</h3>
            <div class="grid sm:grid-cols-2 gap-3 mb-8">
              <div v-for="feature in product.features" :key="feature"
                class="flex items-center gap-2 p-3 rounded-lg bg-gray-50">
                <svg class="w-4 h-4 text-primary flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                  <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                </svg>
                <span class="text-sm text-gray-700">{{ feature }}</span>
              </div>
            </div>
            <h3 class="text-base font-semibold text-gray-900 mb-4">适用企业</h3>
            <p class="text-steel-light text-sm leading-relaxed mb-6">{{ product.target }}</p>
            <router-link to="/trial"
              class="inline-flex items-center gap-2 px-6 py-3 bg-primary text-white font-medium rounded-xl hover:bg-primary-dark transition-colors">
              申请试用
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/></svg>
            </router-link>
          </div>
          <div :class="index % 2 !== 0 ? 'lg:col-start-1' : ''">
            <div class="rounded-2xl p-8 border border-gray-200" :class="product.cardBg">
              <div class="text-center">
                <div class="text-6xl mb-6">{{ product.icon }}</div>
                <h3 class="text-xl font-bold text-gray-900 mb-4">{{ product.name }}</h3>
                <div class="flex flex-wrap justify-center gap-2">
                  <span v-for="kw in product.keywords" :key="kw"
                    class="px-3 py-1 bg-white/80 text-gray-600 text-xs rounded-full border border-gray-200">
                    {{ kw }}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- CTA -->
    <section class="py-16 gradient-hero text-center">
      <div class="max-w-4xl mx-auto px-4">
        <h2 class="text-3xl font-bold text-white mb-6">找到适合您企业的产品</h2>
        <p class="text-white/80 mb-8">联系我们，获取专业的产品选型建议和免费试用体验</p>
        <div class="flex flex-wrap justify-center gap-4">
          <router-link to="/trial" class="px-8 py-4 bg-white text-primary font-semibold rounded-xl hover:bg-gray-100 transition-all shadow-lg">
            申请试用
          </router-link>
          <router-link to="/contact" class="px-8 py-4 border-2 border-white text-white font-semibold rounded-xl hover:bg-white/10 transition-all">
            联系顾问
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
        '@type': 'ItemList',
        name: '普讯软件产品中心',
        description: '南京普讯管理软件有限公司核心产品列表',
        itemListElement: [
          { '@type': 'ListItem', position: 1, name: '钢贸宝ERP', description: '钢铁贸易企业经营数智化ERP管理系统' },
          { '@type': 'ListItem', position: 2, name: '钢企通ERP', description: '钢铁生产制造企业综合管理ERP平台' },
          { '@type': 'ListItem', position: 3, name: '库准 生产MES', description: '仓库与生产制造执行系统' },
          { '@type': 'ListItem', position: 4, name: '智能财务', description: '钢铁行业智能财务管理系统' },
          { '@type': 'ListItem', position: 5, name: '货袋子赋能平台', description: '钢铁产业互联网交易赋能平台' }
        ]
      })
    }
  ]
})

const products = [
  {
    name: '钢贸宝 ERP',
    icon: '📊',
    tag: '核心产品',
    tagClass: 'bg-blue-100 text-blue-700',
    desc: '钢贸宝ERP是专为钢铁贸易企业打造的全流程经营数智化管理系统。覆盖采购、销售、库存、财务、客户管理等全业务链条，帮助钢贸企业从传统经营模式升级为数字化经营模式。已有近3000家钢贸企业选择使用。',
    features: ['采购管理', '销售管理', '库存管理', '财务管理', '客户关系管理', '数据分析报表', '多分公司管理', '销售外出管理', '运营自动化', '移动端支持'],
    target: '适用于各类钢铁贸易企业，从小型贸易商到大型集团钢贸企业，均可通过钢贸宝ERP实现全流程数智化管理。',
    cardBg: 'bg-gradient-to-br from-blue-50 to-indigo-50',
    keywords: ['钢贸ERP', '钢铁贸易管理', '贸易数字化', '进销存管理', '钢贸管理软件']
  },
  {
    name: '钢企通 ERP',
    icon: '🏭',
    tag: '生产制造',
    tagClass: 'bg-green-100 text-green-700',
    desc: '钢企通ERP是面向钢铁生产制造企业的综合管理平台，整合生产计划、仓储物流、销售采购、财务核算等模块，实现生产制造企业的一体化数字管理。',
    features: ['生产计划管理', '仓储物流管理', '销售采购管理', '财务一体化', '集团管控', '智能分析决策', '质量追溯管理', '设备管理'],
    target: '适用于钢铁生产加工企业，特别是需要将生产、销售、仓储、财务进行一体化管理的制造型企业。',
    cardBg: 'bg-gradient-to-br from-green-50 to-emerald-50',
    keywords: ['钢铁生产ERP', '制造业ERP', '生产管理', '一体化管理', '钢铁企业软件']
  },
  {
    name: '库准 生产MES',
    icon: '📦',
    tag: '仓库 · 生产',
    tagClass: 'bg-purple-100 text-purple-700',
    desc: '库准是集仓库管理和生产制造执行于一体的数字化系统。实现仓库作业全流程数字化管理，支持无人值守等先进应用场景，已有100+家标杆企业上线使用。',
    features: ['仓库作业管理', '库存实时监控', '无人值守系统', '出入库管理', '盘点管理', '货位管理', '生产执行管理', '条码/RFID集成'],
    target: '适用于拥有仓库管理需求的钢贸企业和钢铁生产制造企业，尤其适合需要标准化、数字化仓库管理的企业。',
    cardBg: 'bg-gradient-to-br from-purple-50 to-violet-50',
    keywords: ['仓库管理系统', 'WMS系统', '生产MES', '无人值守', '智能仓储']
  },
  {
    name: '智能财务',
    icon: '💰',
    tag: '财务管理',
    tagClass: 'bg-amber-100 text-amber-700',
    desc: '针对钢铁贸易行业特点定制的智能财务管理系统，实现财务核算自动化、对账结算自动化，大幅降低财务工作量和出错率，让财务管理更精准、更高效。',
    features: ['自动对账', '智能核算', '应收应付管理', '资金流水管理', '多维度报表', '税务管理', '成本核算', '预算管理'],
    target: '适用于所有钢铁贸易和生产制造企业的财务部门，特别是业务量大、对账复杂的企业。',
    cardBg: 'bg-gradient-to-br from-amber-50 to-yellow-50',
    keywords: ['钢贸财务系统', '智能对账', '财务自动化', '成本核算', '财务管理软件']
  },
  {
    name: '货袋子赋能平台',
    icon: '🌐',
    tag: '产业互联网',
    tagClass: 'bg-cyan-100 text-cyan-700',
    desc: '货袋子是普讯运营的钢铁产业互联网交易赋能平台，日均线上库存10万吨。帮助钢铁贸易企业实现线上卖货、获取客户、库存共享，打造钢铁行业的产业互联网生态。',
    features: ['线上发布库存', '在线交易', '获客引流', '库存共享', '行情资讯', '物流对接', '金融服务', '数据分析'],
    target: '适用于所有需要拓展线上销售渠道、寻找新客户、进行线上交易的钢铁贸易企业。',
    cardBg: 'bg-gradient-to-br from-cyan-50 to-sky-50',
    keywords: ['钢铁交易平台', '钢材线上交易', '钢铁电商', '钢铁产业互联网', '找钢材']
  }
]
</script>
